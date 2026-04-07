function ensureSubFolder_(parent, subname) {
  if (!subname) return parent;
  const it = parent.getFoldersByName(subname);
  return it.hasNext() ? it.next() : parent.createFolder(subname);
}

function moveFileToFolder_(file, destFolder) {
  const parents = file.getParents();
  while (parents.hasNext()) parents.next().removeFile(file);
  destFolder.addFile(file);
}

function cleanName_(str) {
  return String(str)
    .replace(/[<>:"/\\|?*]/g, '')
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function getPublicUrl_(file) {
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch(e){}
  return file.getUrl();
}

function setCellToPublicUrl_(sheet, row, col, file) {
  const url = getPublicUrl_(file);
  sheet.getRange(row, col).setValue(url);
}

function resolveFileFromCell_(cellValue, baseFolder) {
  if (!cellValue) return null;
  const val = String(cellValue).trim();
  const idMatch = val.match(/[-\w]{25,}/);
  if (idMatch) return DriveApp.getFileById(idMatch[0]);

  if (baseFolder) {
    const it = baseFolder.getFilesByName(val.split('/').pop());
    if (it.hasNext()) return it.next();
  }
  return null;
}
