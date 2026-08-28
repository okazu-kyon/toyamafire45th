# 開発ルール (Antigravity 用)

- **GASへの自動同期とデプロイ**:
  - コードファイル（`code.js` や各種 `.html` ファイルなど）を追加・修正・削除した場合は、その作業ステップの最後（または作業完了時）に、必ず以下のコマンドを順に実行してGoogle Apps Script（クラウド）に変更を反映およびデプロイ（既存のURLを維持した上書き更新）させてください。
    1. コードのアップロード： `cmd /c npx @google/clasp push`
    2. 既存URLでのデプロイ更新： `cmd /c npx @google/clasp deploy -i AKfycbyDQ3h_Vx72YPTgECxYYMUFUTtyQdgFsGtBlhSZrKReuKXFd3yDmcf2N15FnzgE8hgj4Q -d "AIによる自動アップデート"`
  - 反映およびデプロイが正常に完了したことをログで確認し、エラーが発生した場合はユーザーに通知してください。
