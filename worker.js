export default {
  async fetch(request, env) {
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

    let payload;

    try {
      payload = await request.json();
    } catch (error) {
      return jsonResponse({
        ok: false,
        error: 'Invalid JSON'
      }, 400);
    }

    let teamsPayload;

if (payload.eventType === 'newJob') {
  const aiEvaluation = await evaluateJobWithOpenAI(payload, env);
  teamsPayload = buildNewJobTeamsPayload(payload, aiEvaluation);
} else if (payload.eventType === 'sheetUpdated') {
  teamsPayload = buildSheetUpdatedTeamsPayload(payload);
} else {
  return jsonResponse({
    ok: false,
    error: 'Unsupported eventType',
    eventType: payload.eventType || null
  }, 400);
}

    const teamsResponse = await fetch(env.TEAMS_WEBHOOK_URL, {
      method: 'POST',
      headers: {
        'content-type': 'application/json'
      },
      body: JSON.stringify(teamsPayload)
    });

    const teamsText = await teamsResponse.text();

    if (!teamsResponse.ok) {
      return jsonResponse({
        ok: false,
        error: 'Teams webhook failed',
        status: teamsResponse.status,
        body: teamsText
      }, 502);
    }

    return jsonResponse({
      ok: true,
      eventType: payload.eventType
    }, 200);
  }
};

// Simple payload builder
function buildNewJobTeamsPayload(job, aiEvaluation) {
  const projectTitle = safeText(job.projectTitle) || 'タイトル未取得';
  const sheetName = safeText(job.sheetName);
  const jobUrl = safeText(job.jobUrl);
  const sheetUrl = safeText(job.sheetUrl);

  const jobType = sheetName === 'saas-projects' ? 'SaaS案件' : '通常案件';

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
          body: [
            {
              type: 'TextBlock',
              text: '新しい案件が登録されました',
              weight: 'Bolder',
              size: 'Medium',
              wrap: true
            },
            {
              type: 'FactSet',
              facts: [
                {
                  title: '種別',
                  value: jobType
                },
                {
                  title: '案件名',
                  value: projectTitle
                },
                {
                  title: 'AI判定',
                  value: aiEvaluation?.decision || '-'
                },
                {
                  title: 'AIスコア',
                  value: aiEvaluation?.score === null || aiEvaluation?.score === undefined
                    ? '-'
                    : String(aiEvaluation.score)
                },
                {
                  title: 'AI要約',
                  value: aiEvaluation?.summary || '-'
                },
                {
                  title: '主な理由',
                  value: Array.isArray(aiEvaluation?.reasons)
                    ? aiEvaluation.reasons.join('\n')
                    : '-'
                },
                {
                  title: '案件URL',
                  value: jobUrl || '-'
                },
                {
                  title: 'Sheet',
                  value: sheetUrl || '-'
                }
              ]
            }
          ]
        }
      }
    ]
  };
}

async function evaluateJobWithOpenAI(job, env) {
  if (!env.OPENAI_API_KEY) {
    return {
      decision: '未評価',
      score: null,
      summary: 'OPENAI_API_KEY がCloudflareに設定されていません。',
      reasons: []
    };
  }

  const projectTitle = safeText(job.projectTitle);
  const jobUrl = safeText(job.jobUrl);
  const sheetName = safeText(job.sheetName);
  const description = safeText(job.description);
  const budget = safeText(job.budget);
  const deadline = safeText(job.deadline);

  const prompt = `
あなたはLIJ株式会社の案件評価アシスタントです。
以下の案件について、受けるべきかどうかを判定してください。

判定ルール:
- 予算が低すぎる案件は低評価
- 納期が短すぎる案件は注意
- 要件が曖昧な案件は注意
- SaaS案件は継続性がありそうなら高評価
- 技術的に不明点が多い場合は「要確認」
- 最終判断は人間が行うため、あなたは推薦を出すだけ

必ず次のJSONだけを返してください。

{
  "decision": "受けるべき" | "要確認" | "見送り推奨",
  "score": 0,
  "summary": "短い要約",
  "reasons": ["理由1", "理由2", "理由3"]
}

案件情報:
案件名: ${projectTitle}
種別: ${sheetName}
URL: ${jobUrl}
説明: ${description}
予算: ${budget}
納期: ${deadline}
`;

  const response = await fetch('https://api.openai.com/v1/responses', {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${env.OPENAI_API_KEY}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({
      model: 'gpt-4.1-mini',
      input: prompt,
      temperature: 0.2,
      text: {
        format: {
          type: 'json_schema',
          name: 'job_evaluation',
          strict: true,
          schema: {
            type: 'object',
            additionalProperties: false,
            properties: {
              decision: {
                type: 'string',
                enum: ['受けるべき', '要確認', '見送り推奨']
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
            required: ['decision', 'score', 'summary', 'reasons']
          }
        }
      }
})
  });

console.log('OpenAI status:', response.status);


  const data = await response.json();
  console.log('OpenAI response:', JSON.stringify(data).slice(0, 2000));

  if (!response.ok) {
    return {
      decision: 'AI評価エラー',
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
      score: null,
      summary: 'OpenAIの返答をJSONとして解析できませんでした。',
      reasons: [text.slice(0, 500)]
    };
  }
}


// *** Includes Entry Condition, InquiryContent, and Hearing Content
// function buildNewJobTeamsPayload(job) {
//   const projectTitle = safeText(job.projectTitle) || 'タイトル未取得';
//   const sheetName = safeText(job.sheetName);
//   const jobUrl = safeText(job.jobUrl);
//   const sheetUrl = safeText(job.sheetUrl);

//   const entryConditions = truncateText(safeText(job.entryConditions), 900);
//   const inquiryContent = truncateText(safeText(job.inquiryContent), 900);
//   const hearingContent = truncateText(safeText(job.hearingContent), 900);

//   const status = safeText(job.status) || '-';
//   const note = safeText(job.note);

//   const jobType = sheetName === 'saas-projects' ? 'SaaS案件' : '通常案件';

//   const facts = [
//     {
//       title: '種別',
//       value: jobType
//     },
//     {
//       title: '案件名',
//       value: projectTitle
//     },
//     {
//       title: 'Status',
//       value: status
//     },
//     {
//       title: '案件URL',
//       value: jobUrl || '-'
//     },
//     {
//       title: 'Sheet',
//       value: sheetUrl || '-'
//     }
//   ];

//   if (note) {
//     facts.push({
//       title: 'Note',
//       value: note
//     });
//   }

//   return {
//     type: 'message',
//     attachments: [
//       {
//         contentType: 'application/vnd.microsoft.card.adaptive',
//         contentUrl: null,
//         content: {
//           '$schema': 'http://adaptivecards.io/schemas/adaptive-card.json',
//           type: 'AdaptiveCard',
//           version: '1.4',
//           body: [
//             {
//               type: 'TextBlock',
//               text: '新しい案件が登録されました',
//               weight: 'Bolder',
//               size: 'Medium',
//               wrap: true
//             },
//             {
//               type: 'FactSet',
//               facts: facts
//             },
//             {
//               type: 'TextBlock',
//               text: 'エントリー条件',
//               weight: 'Bolder',
//               wrap: true,
//               spacing: 'Medium'
//             },
//             {
//               type: 'TextBlock',
//               text: entryConditions || '-',
//               wrap: true
//             },
//             {
//               type: 'TextBlock',
//               text: 'お問い合わせ時の内容',
//               weight: 'Bolder',
//               wrap: true,
//               spacing: 'Medium'
//             },
//             {
//               type: 'TextBlock',
//               text: inquiryContent || '-',
//               wrap: true
//             },
//             {
//               type: 'TextBlock',
//               text: '発注ナビ担当者のヒアリング内容',
//               weight: 'Bolder',
//               wrap: true,
//               spacing: 'Medium'
//             },
//             {
//               type: 'TextBlock',
//               text: hearingContent || '-',
//               wrap: true
//             }
//           ]
//         }
//       }
//     ]
//   };
// }

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
          body: [
            {
              type: 'TextBlock',
              text: '新しい案件データがGoogleスプレッドシートに登録されました。',
              weight: 'Bolder',
              size: 'Medium',
              wrap: true
            },
            {
              type: 'FactSet',
              facts: [
                {
                  title: '通常案件',
                  value: String(normalRowCount)
                },
                {
                  title: 'SaaS案件',
                  value: String(saasRowCount)
                },
                {
                  title: '合計',
                  value: String(totalRowCount)
                },
                {
                  title: 'Sheet',
                  value: safeText(payload.sheetUrl) || '-'
                }
              ]
            }
          ]
        }
      }
    ]
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

function truncateText(value, maxLength) {
  const text = safeText(value);

  if (text.length <= maxLength) {
    return text;
  }

  return text.substring(0, maxLength) + '...';
}