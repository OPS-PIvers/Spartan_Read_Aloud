import sys

with open('Code.js', 'r') as f:
    content = f.read()

# Fix addNewAssessment
old_block1 = """    // Validate file URL
    const fileId = getFileIdFromUrl(fileUrl);
    // Add new row with file URL and metadata
    const newRow = new Array(13).fill(""); // 13 columns
    newRow[CONSTANTS.COL.PDF_URL] = fileUrl;
    newRow[CONSTANTS.COL.CLASS_NAME] = metadata.className || "";
    newRow[CONSTANTS.COL.INSTRUCTOR] = metadata.instructor || "";
    newRow[CONSTANTS.COL.PASSWORD] = metadata.password || "";
      if (data[i][CONSTANTS.COL.PDF_URL] === fileUrl) {
        return { error: 'This file is already in the database.' };
      }
    }

    // Add new row with file URL and metadata
    const newRow = new Array(13).fill(''); // 13 columns
    newRow[CONSTANTS.COL.PDF_URL] = fileUrl;
    newRow[CONSTANTS.COL.CLASS_NAME] = metadata.className || '';
    newRow[CONSTANTS.COL.INSTRUCTOR] = metadata.instructor || '';
    newRow[CONSTANTS.COL.PASSWORD] = metadata.password || '';
    newRow[CONSTANTS.COL.STUDENT_EMAILS] = parseStudentEmails(metadata.studentEmails || '');
    newRow[CONSTANTS.COL.READ_ALOUD_ENABLED] = metadata.readAloudEnabled !== false;

    sheet.appendRow(newRow);"""

new_block1 = """    // Validate file URL
    const fileId = getFileIdFromUrl(fileUrl);
    if (!fileId) {
      return { error: 'Invalid file URL.' };
    }

    // Check if URL already exists
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][CONSTANTS.COL.PDF_URL] === fileUrl) {
        return { error: 'This file is already in the database.' };
      }
    }

    // Add new row with file URL and metadata
    const newRow = new Array(13).fill(''); // 13 columns
    newRow[CONSTANTS.COL.PDF_URL] = fileUrl;
    newRow[CONSTANTS.COL.CLASS_NAME] = metadata.className || '';
    newRow[CONSTANTS.COL.INSTRUCTOR] = metadata.instructor || '';
    newRow[CONSTANTS.COL.PASSWORD] = metadata.password || '';
    newRow[CONSTANTS.COL.STUDENT_EMAILS] = parseStudentEmails(metadata.studentEmails || '');
    newRow[CONSTANTS.COL.READ_ALOUD_ENABLED] = metadata.readAloudEnabled !== false;

    sheet.appendRow(newRow);"""

content = content.replace(old_block1, new_block1)

old_block2 = """    return {
      success: true,
      rowIndex: rowIndex,
    if (data[i][CONSTANTS.COL.PDF_URL] === fileUrl && !data[i][CONSTANTS.COL.CHUNK_COUNT]) {
      // Found the new row - analyze it
      const fileId = getFileIdFromUrl(fileUrl);
      if (!fileId) continue;

      const readAloudEnabled = data[i][CONSTANTS.COL.READ_ALOUD_ENABLED] !== false;
      Logger.log(`Processing new assessment: ${fileUrl} (Read Aloud: ${readAloudEnabled})`);

      if (!readAloudEnabled) {
/**
 * Processes a specific assessment (runs steps 1 and 2)."""

new_block2 = """    return {
      success: true,
      rowIndex: rowIndex,
      message: 'Assessment added successfully. Processing has been started.'
    };

  } catch (e) {
    Logger.log(`Error in addNewAssessment: ${e.toString()}`);
    return { error: 'Failed to add assessment: ' + e.toString() };
  }
}

/**
 * Processes a specific assessment (runs steps 1 and 2)."""

content = content.replace(old_block2, new_block2)

with open('Code.js', 'w') as f:
    f.write(content)
