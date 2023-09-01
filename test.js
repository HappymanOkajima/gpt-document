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
