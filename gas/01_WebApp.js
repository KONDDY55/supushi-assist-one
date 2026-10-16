function doGet() {
  return HtmlService.createHtmlOutputFromFile('WebApp')
    .setTitle('スプシアシストOne 要件定義プロトタイプ');
}

function listSourceFiles() {
  try {
    const folder = DriveApp.getFolderById(CONFIG.sourceFolderId);
    const files = collectInventory_(folder.getFiles(), CONFIG.maxFiles);
    const folders = collectInventory_(folder.getFolders(), CONFIG.maxFiles);
    return {
      status: 'success',
      folderName: folder.getName(),
      files: files.items,
      folders: folders.items,
      truncated: files.truncated || folders.truncated
    };
  } catch (error) {
    console.error('Drive一覧取得失敗: ' + error.message);
    return { status: 'error', message: '資料一覧を取得できません。実行アカウントのフォルダ閲覧権限とGASの認証を確認してください。' };
  }
}
