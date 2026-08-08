/**
 * SALON — Google Drive backend for the photo gallery
 * -------------------------------------------------
 * 1. Go to script.google.com and create a new project
 *    (this does NOT need to be tied to a Sheet — Drive only).
 * 2. Paste this whole file in, replacing the placeholder code.
 * 3. Deploy > New deployment > Web app
 *      Execute as: Me
 *      Who has access: Anyone
 * 4. Copy the /exec URL into CONFIG.DRIVE_API_URL in gallery.html
 *
 * All photos are stored in a folder named "Salon Gallery" that
 * this script creates automatically in your Drive the first
 * time you upload.
 */

var FOLDER_NAME = 'Salon Gallery';

function doGet(e) {
  var folder = getFolder();
  var files = folder.getFiles();
  var list = [];

  while (files.hasNext()) {
    var f = files.next();
    if (f.isTrashed()) continue;
    list.push(fileToJson(f));
  }

  list.sort(function(a, b) {
    return new Date(b.createdTime) - new Date(a.createdTime);
  });

  return respond(list);
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);

    if (body.action === 'upload') {
      var folder = getFolder();
      var bytes = Utilities.base64Decode(body.data);
      var blob = Utilities.newBlob(bytes, body.mimeType, body.name);
      var file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      return respond({ result: 'success', file: fileToJson(file) });
    }

    if (body.action === 'delete') {
      var file = DriveApp.getFileById(body.id);
      file.setTrashed(true);
      return respond({ result: 'success' });
    }

    return respond({ result: 'error', error: 'Unknown action' });

  } catch (err) {
    return respond({ result: 'error', error: err.message });
  }
}

/**
 * Converts a Drive File into the JSON shape the gallery expects.
 * - url: a resizable thumbnail link, used as the <img> source
 * - shareUrl: the normal Drive viewer link, used for sharing
 */
function fileToJson(f) {
  return {
    id: f.getId(),
    name: f.getName(),
    mimeType: f.getMimeType(),
    url: 'https://drive.google.com/thumbnail?id=' + f.getId() + '&sz=w1200',
    shareUrl: f.getUrl(),
    createdTime: f.getDateCreated().toISOString()
  };
}

function respond(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getFolder() {
  var folders = DriveApp.getFoldersByName(FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(FOLDER_NAME);
}
