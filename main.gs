function downloadImagesToRowFolders() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues();

  // Define which columns contain image URLs (e.g., columns B and D)
  const imageColumns = ['F', 'G', 'H', 'I', 'J', 'K'];
  const columnIndices = imageColumns.map(col => {
    return col.toUpperCase().charCodeAt(0) - 65; // 'A' has ASCII code 65
  });

  // Define the column to use for folder names (e.g., column A)
  const folderNameColumnIndex = 0; // Change this index to the column you want to use for folder names (0 = column A, 1 = column B, etc.)
  const folderNameColumnIndex2 = 1;

  // Define a parent folder in Google Drive where these folders will be created
  const parentFolderName = 'Mamy Poko';
  let parentFolder = getOrCreateFolder(parentFolderName);

  // Iterate through each row starting from the second row (assuming the first row is headers)
  for (let i = 1; i < data.length; i++) {
    const folderName = `${data[i][folderNameColumnIndex]}-${data[i][folderNameColumnIndex2]}`;
    if (folderName && typeof folderName === 'string') {
      try {
        // Create or access a folder for this row
        let rowFolder = getOrCreateFolder(folderName, parentFolder);

        // Upload each image to the row-specific folder if it doesn't already exist
        columnIndices.forEach(colIdx => {
          const imageUrl = data[i][colIdx];
          if (imageUrl && typeof imageUrl === 'string') {
            try {
              const response = UrlFetchApp.fetch(imageUrl);
              const blob = response.getBlob();
              const contentType = blob.getContentType();
              const extension = contentType.split('/')[1];
              const fileName = `Row${i + 1}_Col${colIdx + 1}.${extension}`;

              // Check if the file already exists in the folder
              if (!doesFileExist(rowFolder, fileName)) {
                rowFolder.createFile(blob).setName(fileName);
                Logger.log(`Successfully uploaded to ${folderName}: ${fileName}`);
              } else {
                Logger.log(`File already exists in ${folderName}: ${fileName}, skipping download.`);
              }
            } catch (fetchError) {
              Logger.log(`Failed to fetch image from Row ${i + 1}, Column ${colIdx + 1}: ${fetchError}`);
            }
          } else {
            // Logger.log(`Failed invalid image from Row ${i + 1}, Column ${colIdx + 1}: imageUrl:${imageUrl}`);
          }
        });
      } catch (folderError) {
        Logger.log(`Failed to create/access folder for Row ${i + 1}: ${folderError}`);
      }
    }
  }
}

function getOrCreateFolder(folderName, parentFolder = null) {
  let folders;
  if (parentFolder) {
    folders = parentFolder.getFoldersByName(folderName);
  } else {
    folders = DriveApp.getFoldersByName(folderName);
  }

  if (folders.hasNext()) {
    return folders.next();
  } else {
    if (parentFolder) {
      return parentFolder.createFolder(folderName);
    } else {
      return DriveApp.createFolder(folderName);
    }
  }
}

function doesFileExist(folder, fileName) {
  const files = folder.getFilesByName(fileName);
  return files.hasNext();
}
