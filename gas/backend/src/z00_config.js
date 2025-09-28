// どのファイルでも使える軽量アクセサ
function props_() {
  return PropertiesService.getScriptProperties();
}

/** BOOTSTRAP/TOKEN どちらでも拾う共通のシークレット取得 */
function getSecret_() {
  var s = props_().getProperty('BOOTSTRAP_SECRET') || props_().getProperty('TOKEN_SECRET') || '';
  if (s && typeof s.replace === 'function') s = s.replace(/[\r\n]+$/g, '');
  if (!s) throw new Error('missing secret');
  return s;
}
