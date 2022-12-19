/*
  =============================================================================
  Project Page: https://github.com/cmenon12/contemporary-choir
  Copyright:    (c) 2022 by Christopher Menon
  License:      GNU General Public License, version 3 (GPL-3.0)
                http://www.opensource.org/licenses/gpl-3.0.html
  =============================================================================
 */


/**
 * This will comprehensively remove the user from the folder
 * by iterating through all its files and subfolders recursively.
 */
function main() {

  folderId = "longfolderid";
  email = "myemail@gmail.com";

  recursiveRemoval(folderId, email);

}

function recursiveRemoval(folderId, email) {

  // Get this folder and its contents
  let mainFolder = DriveApp.getFolderById(folderId);
  let folders = mainFolder.getFolders();
  let files = mainFolder.getFiles();

  // Run this recursively on all subfolders
  let folder;
  while (folders.hasNext()) {
    folder = folders.next();
    recursiveRemoval(folder.getId(), email);
  }

  // Iterate over each file
  let file;
  while (files.hasNext()) {
    file = files.next();

    // Revoke permissions for each file
    try {
      file.revokePermissions(email);
      // Logger.log(`Success for ${file.getName()}`);
    } catch (err) {
      if (err.message !== "No such user") {
        Logger.log(`${err} for ${file.getName()}, ${file.getMimeType()}`);
      }
    }

  }

  // Revoke permissions for this folder
  try {
    mainFolder.revokePermissions(email);
    Logger.log(`Success for ${mainFolder.getName()}`);
  } catch (err) {
    if (err.message !== "No such user") {
      Logger.log(`${err} for ${mainFolder.getName()}`);
    }
  }

}
