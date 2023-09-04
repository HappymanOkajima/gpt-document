# GPTドキュメントAdd-on
生成AIの力で文章や画像を生成しドキュメント作成をサポートする、シンプルなGoogleドキュメント用アドオンです。同様の機能を持つ商用アドオンは複数存在しますが、生成AIの典型的なユースケースがどのように動作するのか学びながら、自分でカスタマイズし組織に導入・運用していけるのがこのアドオンのメリットです。

![gpt-document-1](https://github.com/HappymanOkajima/gpt-document/assets/6194144/1b5f1d06-f5d0-4944-b4f1-0b57e762c7f6)

## 何ができるのか
アドオンはサイドバー形式で、対象のドキュメントに対して3つの自動生成機能を利用できます。文章の自動生成機能でドキュメントの大枠を作成し、その後、選択したテキストに対して指示を与えて深掘りしたり、マッチした画像を生成して、仕上げていくと良いでしょう。
### 1.文章の自動生成
ChatGPTのようにプロンプトで指示を与えます。モデル（GPT-3.5 / GPT-4）と創造性を指定可能です。生成された文章はドキュメントの最後に追加されます。

<img width="293" alt="image" src="https://github.com/HappymanOkajima/gpt-document/assets/6194144/03396ca9-92fa-4b41-a5f3-371a009c75f7">

マークダウン形式からGoogleドキュメントネイティブの見出し・表・リストへの自動変換をサポートしているので、例えば次のようなプロンプトを利用することで、見栄えの良いドキュメントを直接生成することができます。これはChatGPTからコピー＆ペーストするのに比べ、圧倒的に効率的です。

> アジャイル開発の内製化を支援するサービスのパンフレット。サービス名は「ワクワク内製化支援」200文字程度。メリットの一覧。競合との比較表。見出しやリスト、表はマークダウンで。

### 2.選択されたテキストに基づく画像生成
選択されたテキストの内容に基づき、DALL-Eまたは、Stable Diffusionで画像を生成します。画像のスタイルを指定することもできます。生成された画像は選択されたテキストの直後に追加されます。

<img width="276" alt="image" src="https://github.com/HappymanOkajima/gpt-document/assets/6194144/66e6c89f-6d59-4526-96cf-9ad6c9ab3515">

画像の質「標準」「高い」は、Stability AIのAPIキーが必要となります。「低い」はOpenAIのAPIキーが利用できます。「高い」を指定すると非常に高品位な画像が生成されますが、相応の時間がかかります。

### 3.選択されたテキストに基づく文章の生成
選択されたテキストの内容に基づき文章を生成します。生成された文章は選択されたテキストの直後に追加されます。

<img width="274" alt="image" src="https://github.com/HappymanOkajima/gpt-document/assets/6194144/70698c9b-f4b9-4eff-9863-9aed9d786b0b">

## インストール
### とりあえず個人で試す
手っ取り早く試すには、Googleドキュメントの「拡張機能」からスクリプトエディタを開き（アドオンは、[Googleドキュメントにバインドされたスクリプト](https://developers.google.com/apps-script/guides/bound?hl=ja)として動作させる必要があります）、ソースコード（code.gsとsidebar.html）をコピー＆ペーストし、後述の必須プロパティとしてAPIキーを設定することです。その後、対象のドキュメントをリロードすると、「拡張機能」メニューに「GPTドキュメント」が表示されます。ただしこの方法では、スクリプトがバインドされたドキュメントでのみ機能が利用できるため、あくまでも個人用・テスト用です。

<img width="454" alt="image" src="https://github.com/HappymanOkajima/gpt-document/assets/6194144/099f1724-9afb-4be7-b268-a1252c481138">

### 組織全体で利用する
組織全体でGoogleドキュメントのアドオンとして利用する場合は、GCP環境を準備し、Google Workspace Marketplace にアドオンとして登録する必要があります。[公式ドキュメント](https://developers.google.com/workspace/marketplace/how-to-publish?hl=ja)などを参考に登録・設定してください。マニフェスト（appsscript.json）も必須となりますので、このリポジトリのものを参考にしてください。

### 必須プロパティ
上記いずれの場合も、以下のスクリプトプロパティを指定する必要があります。スクリプトプロパティはGASエディタから設定してください。
- OPENAI_API_KEY（必須）
- STABILITY_API_KEY（画像の質「標準・高い」を利用したい場合は必須）
 
※これら以外のプロパティは指定しなくてもソースコード中に指定したデフォルトで問題なく動作します。

<img width="889" alt="image" src="https://github.com/HappymanOkajima/gpt-document/assets/6194144/564f61d4-9e52-4cf8-acf7-29bf7123f97b">

## 関連
- [OpenAI 公式サイト](https://platform.openai.com/)。APIキーの取得はこちらから。
- [Stability AI 公式サイト]( https://platform.stability.ai/)。APIキーの取得はこちらから。

