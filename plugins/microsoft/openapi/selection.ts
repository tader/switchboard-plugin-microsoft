/** Bound the reference to useful operations instead of every generated workbook function. */
export function supportedPath(url: string) {
  const workbook = url.match(/\/workbook(.*)$/)?.[1];
  if (workbook !== undefined) return /^(?:|\/(?:createSession|closeSession|refreshSession)|\/application(?:\/calculate)?|\/worksheets(?:\/\{[^/]+\})?(?:\/(?:add|range\([^/]*\)|usedRange\([^/]*\))(?:\/format(?:\/(?:font|fill|borders))?|\/clear)?)?|\/tables(?:\/[^/]+){0,3})$/.test(workbook);
  if (url.startsWith('/sites')) return /^\/sites(?:\/\{site-id\})?(?:\/(?:drive|drives)(?:\/\{drive-id\})?)?$/.test(url) || url === '/sites/root' || url === '/sites/{hostname}:/{site-path}';
  if (url.startsWith('/drives/')) return /^\/drives\/\{drive-id\}(?:|\/root(?:\/children|\/delta\(\))?|\/items\/\{driveItem-id\}(?:\/(?:children|content|copy|createLink|createUploadSession|invite|permissions|preview|restore|thumbnails|versions|checkin|checkout)(?:\/\{[^/]+\})?)?)$/.test(url);
  return true;
}
