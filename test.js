function test_generateTextFromPrompt() {
  generateTextFromPrompt("アジャイル開発の内製化を支援するサービスのパンフレット。サービス名は「ワクワク内製化支援」200文字程度。メリットの一覧。競合との比較表。見出しやリスト、表はマークダウンで。","mid","gpt-3.5-turbo")
}
function test_generateTextFromSelectedText() {
  // テキストを選択した状態で実行する必要がある
  var doc = DocumentApp.getActiveDocument();
  var rangeBuilder = doc.newRange();
  var texts = doc.getBody().getChild(0).asText();
  rangeBuilder.addElement(texts);
  doc.setSelection(rangeBuilder.build());

  generateTextFromSelectedText("100文字に要約して。","mid","gpt-3.5-turbo");
}
function test_generateImageFromSelectedText() {
  // テキストを選択した状態で実行する必要がある
  var doc = DocumentApp.getActiveDocument();
  var rangeBuilder = doc.newRange();
  var texts = doc.getBody().getChild(0).asText();
  rangeBuilder.addElement(texts);
  doc.setSelection(rangeBuilder.build());

  generateImageFromSelectedText("Photograpic","low");
}
function test_addTextToDocument() {
  let text = "今日は**いい天気**ですね";
  addTextToDocument_(text);
  text = "**いい天気**ですね";
  addTextToDocument_(text);
  text = "今日**いい天気**";
  addTextToDocument_(text);
  text = "今日はいい天気";
  addTextToDocument_(text);
}

