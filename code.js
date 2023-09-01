const OPENAI_API_KEY = PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");
const STABILITY_API_KEY = PropertiesService.getScriptProperties().getProperty("STABILITY_API_KEY");
const IMAGE_TEMPERATURE = PropertiesService.getScriptProperties().getProperty("IMAGE_TEMPERATURE") || "0.5";
const IMAGE_MAX_WORDS = PropertiesService.getScriptProperties().getProperty("IMAGE_MAX_WORDS") || "40";
const IMAGE_PROMPT_GEN_MODEL = PropertiesService.getScriptProperties().getProperty("IMAGE_PROMPT_GEN_MODEL") || "gpt-3.5-turbo-16k";

function onOpen() {
  const menu = DocumentApp.getUi().createAddonMenu();
  menu.addItem('GPTドキュメント', 'showSidebar');
  menu.addToUi();
}

function showSidebar() {
  const ver = "(ver:1)";
  const html = HtmlService.createHtmlOutputFromFile('sidebar.html')
      .setTitle('GPTドキュメント' + ver)
      .setWidth(300);
  DocumentApp.getUi().showSidebar(html);
}

/**
 * generateTextFromPrompt
 * 与えられたプロンプトに基づき、OpenAIのChat APIを利用して新しいテキストを生成し、
 * そのテキストをドキュメントに追加します。
 *
 * @param {string} prompt - 新しいテキストの生成のための初期プロンプト。
 * @param {string} creativity - APIの創造性のレベルを示す文字列("low", "medium", "high")。
 * @param {string} model - 使用するOpenAIのモデルの名前。
 * 
 * 副作用:
 * - プロンプトを元にして新しいテキストを生成します。
 * - 生成されたテキストはドキュメントの適切な位置に追加されます。
 *
 * 使用例:
 * generateTextFromPrompt("次の章の冒頭:", "medium", "gpt-3.5-turbo");
 *
 * 注意:
 * - 内部で`callOpenAIChatAPI_`および`addTextToDocument_`関数を使用します。
 */
function generateTextFromPrompt(prompt, creativity, model) {
  const text = callOpenAIChatAPI_(prompt, creativity, model);
  addTextToDocument_(text);
}

/**
 * generateImageFromSelectedText
 * 選択されたテキストを基にして画像を生成し、ドキュメントにその画像を追加します。
 *
 * @param {string} imageCaption - 画像のキャプションまたは説明。
 * @param {string} imageQuality - 画像の品質を示す文字列("low", "mid", "high")。
 * 
 * 副作用:
 * - 選択されたテキストが存在する場合、それを元にして新しい画像を生成します。
 * - 生成された画像はドキュメントの選択されたテキストの後に追加されます。
 *
 * 使用例:
 * generateImageFromSelectedText("美しい夕暮れ", "high");
 *
 * 注意:
 * - この関数は、ドキュメント内の選択されたテキストを前提として動作します。
 * - 画像の生成は、`generateImageFromDallE_`や`generateImageFromStableDiffusion_`関数を利用します。
 * - `getSelectedText_`, `getImagePrompt_`, `insertImageAfterSelectedText_`などの関数も内部で使用します。
 */
function generateImageFromSelectedText(imageCaption, imageQuality) {
  const selectedText = getSelectedText_();
  if (!selectedText) {
    return;
  }
  const prompt = getImagePrompt_(selectedText);
  let image;
  if (imageQuality === "low") {    
    image = generateImageFromDallE_(prompt,imageCaption);
  } else if (imageQuality === "mid") {
    image = generateImageFromStableDiffusion_(prompt, imageCaption, false);
  } else {
    image = generateImageFromStableDiffusion_(prompt, imageCaption, true);    
  }
  insertImageAfterSelectedText_(image);
}

/**
 * generateTextFromSelectedText
 * 選択されたテキストを使用して新しいテキストを生成します。
 * 新しいテキストは、OpenAIのChat APIを利用して生成されます。
 *
 * @param {string} prompt - 新しいテキストの生成のための初期プロンプト。
 * @param {string} creativity - APIの創造性のレベルを示す文字列("low", "medium", "high")。
 * @param {string} model - 使用するOpenAIのモデルの名前。
 * 
 * 副作用:
 * - 選択されたテキストがある場合、それをプロンプトとして使用して新しいテキストを生成します。
 * - 新しいテキストはドキュメントの適切な位置に追加されます。
 *
 * 使用例:
 * generateTextFromSelectedText("要約してください:", "high", "gpt-3.5-turbo");
 *
 * 注意: 
 * - この関数は、ドキュメント内の選択されたテキストを前提として動作します。
 * - 内部で`getSelectedText_`, `callOpenAIChatAPI_`, `getParentParagraph_`, `addTextToDocument_`などの関数を使用します。
 */
function generateTextFromSelectedText(prompt, creativity, model) {
  const selectedText = getSelectedText_();
  if (!selectedText) {
    return;
  }
  const modifiedPrompt = prompt + "\n" + selectedText;
  const newText = callOpenAIChatAPI_(modifiedPrompt, creativity, model);

  let parentParagraph = getParentParagraph_();
  if (parentParagraph.getType() === DocumentApp.ElementType.TABLE_CELL) {
    parentParagraph = parentParagraph.getParent().getParent().asTable();
  }
  addTextToDocument_(newText, parentParagraph);
}

// 以降プライベートな関数

/**
 * getParentParagraph_
 * アクティブなGoogleドキュメントから、現在選択されている要素の親の段落またはリストアイテムを取得します。
 * 選択されている要素がテーブルセル内の場合、そのテーブルセルが返されます。
 *
 * @return {DocumentApp.Paragraph | DocumentApp.ListItem | DocumentApp.TableCell} 
 *         親の段落、リストアイテム、またはテーブルセル。選択がない場合や適切な親が見つからない場合はnull。
 *
 * 使用例:
 * const parent = getParentParagraph_();
 *
 * 注意:
 * - 選択されたテキストがテーブルセル内の場合、そのテーブルセルが返されます。
 * - 選択されたテキストがリストアイテムの場合、そのリストアイテムが返されます。
 * - それ以外の場合、選択されたテキストの親の段落が返されます。
 */
function getParentParagraph_() {
  let parentParagraph = null;
  const doc = DocumentApp.getActiveDocument();
  const selection = doc.getSelection();

  if (selection) {
    const elements = selection.getSelectedElements();
    const lastElement = elements[elements.length - 1];
    const elementType = lastElement.getElement().getType();
    if (elementType === DocumentApp.ElementType.TEXT) {
      const textElement = lastElement.getElement().asText();
      const parentElement = textElement.getParent();
      if (parentElement.getType() === DocumentApp.ElementType.LIST_ITEM) {
        parentParagraph = parentElement.asListItem();
      } else {
        if (parentElement.getParent().getType() === DocumentApp.ElementType.TABLE_CELL) {
          //parentParagraph = parentElement.getParent().getParent().getParent().asTable();
          parentParagraph = parentElement.getParent().asTableCell();          
        } else {
          parentParagraph = parentElement.asParagraph();
        }
      }
    } else if (elementType === DocumentApp.ElementType.PARAGRAPH) {
      parentParagraph = lastElement.getElement().asParagraph();
    } else if (elementType === DocumentApp.ElementType.LIST_ITEM) {
      parentParagraph = lastElement.getElement().asListItem();
    }
  }
  return parentParagraph;
}

/**
 * addTextToDocument_
 * アクティブなGoogleドキュメントにテキストを追加します。指定されたテキストはMarkdown形式であると想定し、
 * 適切にヘッダ、リスト、テーブルなどの形式で挿入されます。
 *
 * @param {string} text - Googleドキュメントに追加するテキスト。Markdown形式であると想定。
 * @param {DocumentApp.Paragraph | DocumentApp.ListItem | DocumentApp.TableCell} [targetParagraph]
 *        テキストを挿入する位置の基準となる段落、リストアイテム、またはテーブルセル。
 *        指定しない場合は、ドキュメントの末尾に追加されます。
 *
 * 注意事項:
 * - ヘッダは '#' で始まる行として認識され、適切な見出しレベルでドキュメントに追加されます。
 * - リストアイテムは '*'、'+'、または '-' で始まる行として認識されます。
 * - テーブルは '|' を含む行として認識され、Markdown形式のテーブルとして扱われます。
 */
function addTextToDocument_(text, targetParagraph) {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();

  let paragraphIndex;
  if (targetParagraph) {
    paragraphIndex = body.getChildIndex(targetParagraph) + 1;
  } else {
    paragraphIndex = body.getNumChildren() - 1;
  }

  const paragraphs = text.split('\n');
  let inTable = false;
  let tableData = [];

  for (let i = 0; i < paragraphs.length; i++) {
    const paragraph = paragraphs[i];

    // Check for headings
    if (paragraph.startsWith('#')) {
      const headingLevel = paragraph.lastIndexOf('#') + 1;
      const headingText = paragraph.slice(headingLevel).trim();

      // Append heading
      const heading = body.insertParagraph(paragraphIndex++, headingText);
      heading.setHeading(DocumentApp.ParagraphHeading["HEADING" + headingLevel]);
    }
    // Check for unordered lists
    else if (paragraph.startsWith('*') || paragraph.startsWith('+') || paragraph.startsWith('-')) {
      const listItem = body.insertListItem(paragraphIndex++, paragraph.slice(1).trim());
      listItem.setGlyphType(DocumentApp.GlyphType.BULLET);
    }
    // Check for tables
    else if (paragraph.includes('|')) {
      if (!inTable) {
        inTable = true;
      }

      // Split cells and remove outer cells if empty
      const cells = paragraph.split('|').map(cell => cell.trim());
      if (cells[0] === '') cells.shift();
      if (cells[cells.length - 1] === '') cells.pop();

      // Skip separator line in markdown table
      if (cells.some(cell => cell.includes('-'))) {
        continue;
      }

      tableData.push(cells);
    }
    // If not in table anymore, append the table and reset table data
    else if (inTable) {
      appendTableToDocument_(body, tableData, paragraphIndex++);
      tableData = [];
      inTable = false;
    }
    // Append plain paragraph
    else {
      body.insertParagraph(paragraphIndex++, paragraph);
    }
  }

  // If still in table at the end of the loop, append the table
  if (inTable) {
    appendTableToDocument_(body, tableData, paragraphIndex++);
  }
}

/**
 * appendTableToDocument_
 * Googleドキュメントの指定された位置にテーブルを追加します。テーブルデータは2次元配列として提供され、
 * 各配列のエントリはテーブルのセルの内容となります。
 *
 * @param {DocumentApp.Body} body - Googleドキュメントのボディオブジェクト。
 * @param {Array<Array<string>>} tableData - テーブルのデータ。2次元配列で、各エントリはセルの内容を示します。
 * @param {number} paragraphIndex - テーブルを挿入する位置の基準となる段落のインデックス。
 *
 * 注意事項:
 * - 空のテーブルデータが提供された場合、テーブルは追加されません。
 * - 指定された位置にテーブルが追加され、その後のコンテンツは下に移動します。
 */
function appendTableToDocument_(body, tableData, paragraphIndex) {
  if (tableData.length === 0 || tableData[0].length === 0) {
    return; // Do not append an empty table
  }

  const numRows = tableData.length;
  const numColumns = tableData[0].length;

  // Create and append table
  const table = body.insertTable(paragraphIndex);
  for (let i = 0; i < numRows; i++) {
    const row = i === 0 ? table.appendTableRow() : table.appendTableRow();
    for (let j = 0; j < numColumns; j++) {
      const cell = row.appendTableCell();
      cell.setText(tableData[i][j]);
    }
  }
}

/**
 * getSelectedText_
 * 現在のGoogleドキュメントで選択されているテキストを取得します。テキスト、段落、またはリストアイテムが
 * 選択されている場合、そのテキストの内容が返されます。複数のエリアが選択されている場合、すべての選択エリアの
 * テキストが連結されて返されます。
 *
 * @return {string|null} 選択されたテキストの内容。何も選択されていない場合はnullを返します。
 *
 * 注意事項:
 * - この関数は現在のアクティブなドキュメントの選択範囲にのみ動作します。
 * - 選択範囲が複数ある場合、それらは連結されて一つの文字列として返されます。
 * - 段落が選択されている場合、各段落の末尾には改行("\n")が追加されます。
 */
function getSelectedText_() {
  const doc = DocumentApp.getActiveDocument();
  const selection = doc.getSelection();

  if (selection) {
    const elements = selection.getSelectedElements();
    let selectedText = "";

    elements.forEach(function(element) {
      const elementType = element.getElement().getType();

      if (elementType === DocumentApp.ElementType.TEXT || elementType === DocumentApp.ElementType.PARAGRAPH || elementType === DocumentApp.ElementType.LIST_ITEM) {
        const textElement = element.getElement().asText();
        const startOffset = element.getStartOffset();
        const endOffset = element.getEndOffsetInclusive();

        if (startOffset !== -1 && endOffset !== -1) {
          selectedText += textElement.getText().substring(startOffset, endOffset + 1);
        } else if (startOffset === -1 && endOffset !== -1) {
          selectedText += textElement.getText().substring(0, endOffset + 1);
        } else if (startOffset !== -1 && endOffset === -1) {
          selectedText += textElement.getText().substring(startOffset);
        } else {
          selectedText += textElement.getText();
        }

        if (elementType === DocumentApp.ElementType.PARAGRAPH) {
          selectedText += "\n";
        }
      }
    });

    return selectedText;
  } else {
    return null;
  }
}

/**
 * insertImageAfterSelectedText_
 * 現在のGoogleドキュメントで選択されているテキストの後ろに画像を挿入します。選択範囲がテーブルセル内にある場合、
 * 画像はそのセルの最後に挿入されます。選択範囲がテーブルセル外にある場合、画像はその段落の最後に挿入されます。
 * 挿入される画像のサイズはデフォルトで300x300に設定されます。
 *
 * @param {Blob} blob - 挿入する画像のBlobオブジェクト。
 *
 * 注意事項:
 * - 選択範囲がない場合、関数は何も行わず終了します。
 * - 画像のサイズはデフォルトで300x300に固定されています。必要に応じてこの値を変更することができます。
 */
function insertImageAfterSelectedText_(blob) {
  const doc = DocumentApp.getActiveDocument();
  const selection = doc.getSelection();
  const imageBlob = blob;
  if (!selection) {
    return;
  }

  const parentParagraph = getParentParagraph_();
  let inlineImage;
  if (parentParagraph.getType() === DocumentApp.ElementType.TABLE_CELL) {
    inlineImage = parentParagraph.appendImage(imageBlob);
  } else {
    inlineImage = parentParagraph.appendInlineImage(imageBlob);
  }
  inlineImage.setWidth(300);
  inlineImage.setHeight(300);

}

/**
 * callOpenAIChatAPI_
 * OpenAIのChat APIを利用して、与えられたプロンプトに基づく返答を取得します。
 *
 * @param {string} prompt - 返答を取得するための初期プロンプト。
 * @param {string} creativity - APIの創造性のレベルを示す文字列("low", "medium", "high")。
 * @param {string} model - 使用するOpenAIのモデルの名前。
 * 
 * @returns {string} APIからの返答。
 *
 * 注意:
 * - この関数は、`OPENAI_API_KEY`の設定を前提としています。
 * - 返答の創造性は、`creativity`パラメータによって調整されます。
 *   "low"はテンプレート的な返答を、"high"はより創造的な返答を示します。
 *
 * 使用例:
 * const reply = callOpenAIChatAPI_("こんにちは", "medium", "gpt-3.5-turbo");
 * console.log(reply);
 */
function callOpenAIChatAPI_(prompt, creativity, model) {
  let temperature = 0.5;
  if (creativity === "low") {
    temperature = 0.0;
  } else if (creativity === "high") {
    temperature = 1.0;
  }

  const url = "https://api.openai.com/v1/chat/completions";
  const data = {
    model: model,
    messages: [
      { role: "user", content: prompt }
    ],
    temperature: temperature
  };

  const options = {
    method: "post",
    headers: {
      Authorization: "Bearer " + OPENAI_API_KEY,
      "Content-Type": "application/json"
    },
    payload: JSON.stringify(data),
    muteHttpExceptions: true
  };

  log_("【REQ_1】" + JSON.stringify(data));

  const response = UrlFetchApp.fetch(url, options);
  const jsonResponse = JSON.parse(response.getContentText());
  const content = jsonResponse.choices[0].message.content.trim();

  log_("【RES_1】" + content);
  return content;
}

/**
 * getImagePrompt_
 * OpenAIのAPIを利用して、与えられたテキストに基づく画像生成のための詳細なプロンプトを取得します。
 *
 * @param {string} targetText - 画像生成のプロンプトとして変換されるテキスト。
 * 
 * @returns {string} 画像生成AI向けの詳細なプロンプト。
 *
 * 内部関数:
 * - getPositivePrompt: 画像生成のための基本的なプロンプト文を生成する。
 *
 * 使用例:
 * const detailedPrompt = getImagePrompt_("夕日が沈むビーチの風景");
 * console.log(detailedPrompt);
 *
 * 注意: この関数の実行には、OPENAI_API_KEYおよびIMAGE_PROMPT_GEN_MODELの設定が必要です。
 * また、IMAGE_MAX_WORDSとIMAGE_TEMPERATUREの値も考慮に入れられます。
 */
function getImagePrompt_(targetText) {
  const url = "https://api.openai.com/v1/chat/completions";

  const getPositivePrompt = function() {
    const prompt = `Please imagine a common scene from the given text and express it as a detailed prompt in english for image generation AI.\n
  rule:\n- begin  a sentence with  "The prompt is:".\n- without using bullet points.\n- in the third person.\n- ${IMAGE_MAX_WORDS} words.\ntext :`;
    return prompt;
  }

  const data = {
    model: IMAGE_PROMPT_GEN_MODEL,
    messages: [
      {
        role: "system",
        content: getPositivePrompt()
      },
      { role: "user", content: targetText || 'The message "error" in white wall.' }
    ],
    max_tokens: Math.round(Number.parseInt(IMAGE_MAX_WORDS) * 3),
    temperature: Number.parseFloat(IMAGE_TEMPERATURE)
  };

  const options = {
    method: "post",
    headers: {
      Authorization: "Bearer " + OPENAI_API_KEY,
      "Content-Type": "application/json"
    },
    payload: JSON.stringify(data),
    muteHttpExceptions: true
  };

  log_("【REQ_2】" + JSON.stringify(data));

  const response = UrlFetchApp.fetch(url, options);
  const jsonResponse = JSON.parse(response.getContentText());
  const content = jsonResponse.choices[0].message.content.trim();

  log_("【RES_2】" + content);

  return content;
}

/**
 * generateImageFromDallE_
 * OpenAIのDALL·E APIを利用して、テキストプロンプトから画像を生成します。
 *
 * @param {string} prompt - 画像生成の主要なプロンプト。
 * @param {string} imageCaption - 画像に関連するキャプション。
 * 
 * @returns {Blob} 生成された画像のBlobオブジェクト。
 *
 * 使用例:
 * const imageBlob = generateImageFromDallE_("猫", "Anime");
 * DriveApp.createFile(imageBlob);
 *
 * 注意: この関数の実行には、OPENAI_API_KEYの設定が必要です。
 */
function generateImageFromDallE_(prompt, imageCaption) {
  // 必要な情報をセットアップします
  const apiUrl = "https://api.openai.com/v1/images/generations";
  const model = "image-alpha-001";
  const size = "512x512";
  const responseFormat = "url";

  // APIリクエストのヘッダーとペイロードを設定します
  const options = {
    method: "post",
    headers: {
      "Content-Type": "application/json",
      "Authorization": "Bearer " + OPENAI_API_KEY
    },
    payload: JSON.stringify({
      model: model,
      prompt: prompt + " " + imageCaption,
      size: size,
      response_format: responseFormat
    })
  };

  // APIリクエストを実行し、結果を取得します
  const response = UrlFetchApp.fetch(apiUrl, options);
  const jsonResponse = JSON.parse(response.getContentText());
  const imageUrl = jsonResponse.data[0].url;
  return UrlFetchApp.fetch(imageUrl).getBlob();
}

/**
 * generateImageFromStableDiffusion_
 * Stability APIを利用して、テキストプロンプトから画像を生成します。
 *
 * @param {string} prompt - 画像生成の主要なプロンプト。
 * @param {string} imageCaption - 画像に関連するキャプション。
 * @param {boolean} highQuality - 高品質の画像を生成する場合はtrue。それ以外の場合はfalse。
 * 
 * @returns {Blob} 生成された画像のBlobオブジェクト。
 *
 * @throws {Error} Stability APIキーが見つからない場合、エラーをスローします。
 * 
 * 使用例:
 * const imageBlob = generateImageFromStableDiffusion_("猫", "Anime", true);
 * DriveApp.createFile(imageBlob);
 *
 * 注意: この関数は内部的な使用のためのものであり、STABILITY_API_KEYの設定が前提となります。
 */
function generateImageFromStableDiffusion_(prompt, imageCaption, highQuality) {
  const api_key = STABILITY_API_KEY;
  const engine_id = highQuality ? "stable-diffusion-xl-1024-v1-0" : "stable-diffusion-xl-beta-v2-2-2";
  const size = highQuality ? 1024: 512;
  const api_host = "https://api.stability.ai";

  if (api_key === null) {
    throw new Error("Missing Stability API key.");
  }

  const text = `${prompt} , ${imageCaption}`;

  const response = UrlFetchApp.fetch(`${api_host}/v1/generation/${engine_id}/text-to-image`, {
    method: "post",
    headers: {
      "Content-Type": "application/json",
      "Accept": "application/json",
      "Authorization": `Bearer ${api_key}`,
    },
    payload: JSON.stringify({
      text_prompts: [
        {
          text: text,
          weight:0.9
        },
        {
          text: "lowres, bad anatomy, bad hands, text, error, missing fingers, extra digit, fewer digits, cropped, worst quality, low quality, normal quality, jpeg artifacts, signature, watermark, username, blurry",
          weight:-1.0
        }
      ],
      cfg_scale: 20,
      clip_guidance_preset: "FAST_BLUE",
      height: size,
      width: size,
      samples: 1,
      steps: 50,
      sampler: "K_EULER_ANCESTRAL",
      seed: 0 
    }),
  });

  const data = JSON.parse(response.getContentText());
  let blob;
  for (let i = 0; i < data.artifacts.length; i++) {
    const image = data.artifacts[i];
    const decodedImage = Utilities.base64Decode(image.base64);
    blob = Utilities.newBlob(decodedImage, "image/png", `v1_txt2img_${i}.png`);
  }
  return blob;
}

function log_(message) {
  const log = `${new Date().toLocaleString()}: ${message}`;
  console.log(log);
}
