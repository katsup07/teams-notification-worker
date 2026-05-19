import hachuNaviPrompt from './prompt-assets/hachunavi-judgement-prompt.txt';
import lijBasicInfo from './prompt-assets/lij-basic-info.txt';
import techCriteria from './prompt-assets/tech-market-value-criteria.txt';
import mismatchCheck from './prompt-assets/mismatch-check.txt';
import entryTemplate from './prompt-assets/entry-template.txt';

const OPENAI_EVALUATION_SCHEMA = {
  type: 'object',
  additionalProperties: false,
  properties: {
    decision: {
      type: 'string',
      enum: ['受けるべき', '要確認', '見送り推奨']
    },
    rank: {
      type: 'string',
      enum: ['S', 'A', 'B', 'C']
    },
    score: {
      type: 'number',
      minimum: 0,
      maximum: 100
    },
    summary: {
      type: 'string'
    },
    reasons: {
      type: 'array',
      items: {
        type: 'string'
      }
    }
  },
  required: ['decision', 'rank', 'score', 'summary', 'reasons']
};

// Main Function
export default {
  async fetch(request, env) {
    const validationError = validateRequest(request, env);
    if (validationError) return validationError;

    const { payload, errorResponse } = await parseJsonRequest(request);
    if (errorResponse) return errorResponse;

    const teamsPayloadResult = await buildTeamsPayloadForEvent(payload, env);
    if (teamsPayloadResult.errorResponse) return teamsPayloadResult.errorResponse;

    const teamsError = await postToTeams(env.TEAMS_WEBHOOK_URL, teamsPayloadResult.payload);
    if (teamsError) return teamsError;

    return jsonResponse({
      ok: true,
      eventType: payload.eventType
    }, 200);
  }
};

function validateRequest(request, env) {
  if (request.method !== 'POST') {
    return jsonResponse({
      ok: false,
      error: 'Method not allowed'
    }, 405);
  }

  const apiKey = String(request.headers.get('x-api-key') || '').trim();
  const expectedApiKey = String(env.APPS_SCRIPT_API_KEY || '').trim();

  if (!apiKey || apiKey !== expectedApiKey) {
    return jsonResponse({
      ok: false,
      error: 'Unauthorized',
      hasReceivedApiKey: Boolean(apiKey),
      hasExpectedApiKey: Boolean(expectedApiKey),
      receivedApiKeyLength: apiKey.length,
      expectedApiKeyLength: expectedApiKey.length
    }, 401);
  }

  if (!env.TEAMS_WEBHOOK_URL) {
    return jsonResponse({
      ok: false,
      error: 'TEAMS_WEBHOOK_URL is not set in Cloudflare Worker variables'
    }, 500);
  }

  return null;
}

async function parseJsonRequest(request) {
  try {
    return {
      payload: await request.json(),
      errorResponse: null
    };
  } catch (error) {
    return {
      payload: null,
      errorResponse: jsonResponse({
        ok: false,
        error: 'Invalid JSON'
      }, 400)
    };
  }
}

async function buildTeamsPayloadForEvent(payload, env) {
  if (payload.eventType === 'newJob') {
    const aiEvaluation = await evaluateJobWithOpenAI(payload, env);

    return {
      payload: buildNewJobTeamsPayload(payload, aiEvaluation),
      errorResponse: null
    };
  }

  if (payload.eventType === 'sheetUpdated') {
    return {
      payload: buildSheetUpdatedTeamsPayload(payload),
      errorResponse: null
    };
  }

  return {
    payload: null,
    errorResponse: jsonResponse({
      ok: false,
      error: 'Unsupported eventType',
      eventType: payload.eventType || null
    }, 400)
  };
}

async function postToTeams(webhookUrl, payload) {
  const teamsResponse = await fetch(webhookUrl, {
    method: 'POST',
    headers: {
      'content-type': 'application/json'
    },
    body: JSON.stringify(payload)
  });

  const teamsText = await teamsResponse.text();

  if (teamsResponse.ok) {
    return null;
  }

  return jsonResponse({
    ok: false,
    error: 'Teams webhook failed',
    status: teamsResponse.status,
    body: teamsText
  }, 502);
}

function buildNewJobTeamsPayload(job, aiEvaluation) {
  const projectTitle = safeText(job.projectTitle) || 'タイトル未取得';
  const sheetName = safeText(job.sheetName);
  const jobUrl = safeText(job.jobUrl);
  const sheetUrl = safeText(job.sheetUrl);

  const jobType = sheetName === 'saas-projects' ? 'SaaS案件' : '開発・制作案件';
  const rankWithScore = formatRankWithScore(aiEvaluation);

  return buildAdaptiveCardMessage([
    headingBlock('新しい案件が登録されました'),
    factSet([
      fact('種別', jobType),
      fact('案件名', projectTitle),
      fact('AI判定', aiEvaluation?.decision || '-'),
      fact('ランク', rankWithScore),
      fact('AI要約', aiEvaluation?.summary || '-'),
      fact('主な理由', formatReasons(aiEvaluation?.reasons)),
      fact('案件URL', jobUrl || '-'),
      fact('Sheet', sheetUrl || '-')
    ])
  ]);
}

async function evaluateJobWithOpenAI(job, env) {
  if (!env.OPENAI_API_KEY) {
    return missingOpenAIKeyEvaluation();
  }

  const prompt = buildOpenAIPrompt(normalizeJobForEvaluation(job));

  const response = await fetch('https://api.openai.com/v1/responses', {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${env.OPENAI_API_KEY}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify(buildOpenAIRequestBody(prompt))
  });

  console.log('OpenAI status:', response.status);

  const data = await response.json();
  console.log('OpenAI response:', JSON.stringify(data).slice(0, 2000));

  return parseOpenAIEvaluationResponse(response, data);
}

function missingOpenAIKeyEvaluation() {
  return {
    decision: '未評価',
    rank: null,
    score: null,
    summary: 'OPENAI_API_KEY がCloudflareに設定されていません。',
    reasons: []
  };
}

function normalizeJobForEvaluation(job) {
  return {
    projectTitle: safeText(job.projectTitle),
    jobUrl: safeText(job.jobUrl),
    sheetName: safeText(job.sheetName),
    scrapedAt: firstText(job, ['scrapedAt', 'scraped_at']),
    status: firstText(job, ['status']),
    note: firstText(job, ['note']),
    entryConditions: firstText(job, ['entryConditions', 'エントリー条件']),
    inquiryContent: firstText(job, ['inquiryContent', 'お問い合わせ時の内容', 'description']),
    hearingContent: firstText(job, ['hearingContent', '発注ナビ担当者のヒアリング内容']),
    budget: firstText(job, ['budget', '予算']),
    deadline: firstText(job, ['deadline', '納期'])
  };
}

function buildOpenAIRequestBody(prompt) {
  return {
    model: 'gpt-5.4-mini',
    input: prompt,
    temperature: 0.2,
    text: {
      format: {
        type: 'json_schema',
        name: 'job_evaluation',
        strict: true,
        schema: OPENAI_EVALUATION_SCHEMA
      }
    }
  };
}

function parseOpenAIEvaluationResponse(response, data) {
  if (!response.ok) {
    const text = extractOpenAIOutputText(data);

    return {
      decision: 'AI評価エラー',
      rank: null,
      score: null,
      summary: 'OpenAIの返答をJSONとして解析できませんでした。',
      reasons: [
        text ? text.slice(0, 500) : 'OpenAI output_text was empty'
      ]
    };
  }

  const text = extractOpenAIOutputText(data);
  console.log('OpenAI raw text:', text);

  try {
    return JSON.parse(text);
  } catch (error) {
    return {
      decision: 'AI評価エラー',
      rank: null,
      score: null,
      summary: 'OpenAIの返答をJSONとして解析できませんでした。',
      reasons: [text.slice(0, 500)]
    };
  }
}

function buildOpenAIPrompt(job) {
  return `
${hachuNaviPrompt}

---
## LIJ基本情報
${lijBasicInfo}

---
## 技術市場価値評価基準
${techCriteria}

---
## 判定ズレチェック
${mismatchCheck}

---
## エントリー文面
${entryTemplate}

---
## Worker出力制約

このCloudflare WorkerではTeams通知用に構造化JSONだけを受け取る。
上記プロンプト内にエントリー文面ドラフトの指示があっても、このAPI応答では本文ドラフトを出力しない。

判定ランクを以下のJSON形式に変換して返すこと。
- rankには必ず元の判定ランク（"S"、"A"、"B"、"C" のいずれか）を入れる
- SまたはA: decision = "受けるべき"
- B: decision = "要確認"
- C: decision = "見送り推奨"
- scoreは0から100。S=90以上、A=75以上、B=45から74、C=44以下を目安にする

必ず次のJSONだけを返すこと。Markdownや説明文をJSONの外に出さないこと。

{
  "decision": "受けるべき" | "要確認" | "見送り推奨",
  "rank": "S" | "A" | "B" | "C",
  "score": 0,
  "summary": "短い要約",
  "reasons": ["理由1", "理由2", "理由3"]
}

---
## 発注ナビ案件情報

job_url: ${job.jobUrl}
project_title: ${job.projectTitle}
sheetName: ${job.sheetName}
scraped_at: ${job.scrapedAt}
status: ${job.status}
note: ${job.note}

エントリー条件:
${job.entryConditions || '-'}

お問い合わせ時の内容:
${job.inquiryContent || '-'}

発注ナビ担当者のヒアリング内容:
${job.hearingContent || '-'}

予算:
${job.budget || '-'}

納期:
${job.deadline || '-'}
`;
}

function extractOpenAIOutputText(data) {
  if (data.output_text) {
    return data.output_text;
  }

  if (!Array.isArray(data.output)) {
    return '';
  }

  const parts = [];

  for (const item of data.output) {
    if (!Array.isArray(item.content)) continue;

    for (const content of item.content) {
      if (content.type === 'output_text' && content.text) {
        parts.push(content.text);
      }
    }
  }

  return parts.join('\n').trim();
}

function buildSheetUpdatedTeamsPayload(payload) {
  const normalRowCount = Number(payload.normalRowCount || 0);
  const saasRowCount = Number(payload.saasRowCount || 0);
  const totalRowCount = normalRowCount + saasRowCount;

  return buildAdaptiveCardMessage([
    headingBlock('新しい案件データがGoogleスプレッドシートに登録されました。'),
    factSet([
      fact('開発・制作案件', String(normalRowCount)),
      fact('SaaS案件', String(saasRowCount)),
      fact('合計', String(totalRowCount)),
      fact('Sheet', safeText(payload.sheetUrl) || '-')
    ])
  ]);
}

function buildAdaptiveCardMessage(body) {
  return {
    type: 'message',
    attachments: [
      {
        contentType: 'application/vnd.microsoft.card.adaptive',
        contentUrl: null,
        content: {
          '$schema': 'http://adaptivecards.io/schemas/adaptive-card.json',
          type: 'AdaptiveCard',
          version: '1.4',
          body
        }
      }
    ]
  };
}

function headingBlock(text) {
  return {
    type: 'TextBlock',
    text,
    weight: 'Bolder',
    size: 'Medium',
    wrap: true
  };
}

function factSet(facts) {
  return {
    type: 'FactSet',
    facts
  };
}

function fact(title, value) {
  return {
    title,
    value
  };
}

function jsonResponse(body, status) {
  return new Response(JSON.stringify(body), {
    status: status,
    headers: {
      'content-type': 'application/json'
    }
  });
}

function safeText(value) {
  if (value === null || value === undefined) {
    return '';
  }

  return String(value);
}

function firstText(source, keys) {
  for (const key of keys) {
    const value = source?.[key];

    if (value !== null && value !== undefined && String(value).trim() !== '') {
      return String(value);
    }
  }

  return '';
}

function formatRankWithScore(aiEvaluation) {
  const rank = safeText(aiEvaluation?.rank);
  const score = aiEvaluation?.score;

  if (!rank && (score === null || score === undefined)) {
    return '-';
  }

  if (!rank) {
    return `${score}点`;
  }

  if (score === null || score === undefined) {
    return rank;
  }

  return `${rank}（${score}点）`;
}

function formatReasons(reasons) {
  return Array.isArray(reasons) ? reasons.join('\n') : '-';
}
