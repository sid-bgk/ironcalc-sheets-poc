const { google } = require('googleapis');

const auth = new google.auth.GoogleAuth({
  keyFile: "credentials.json",
  scopes: ["https://www.googleapis.com/auth/drive"]
});

const drive = google.drive({ version: "v3", auth });

const cleanup = async () => {
  // First, empty trash
  console.log("Emptying service account's trash...");
  try {
    await drive.files.emptyTrash();
    console.log("✓ Trash emptied\n");
  } catch (err) {
    console.log(`✗ Could not empty trash: ${err.message}\n`);
  }

  console.log("Fetching ALL files owned by service account...\n");

  // List ALL files (not just spreadsheets)
  const res = await drive.files.list({
    pageSize: 500,
    fields: "files(id, name, mimeType, size, createdTime)",
    q: "'me' in owners"  // Only files owned by service account
  });

  const files = res.data.files;

  if (!files || files.length === 0) {
    console.log("No files found. Storage might be in Trash.");
    console.log("\nTry emptying trash at: https://drive.google.com/drive/trash");
    return;
  }

  console.log(`Found ${files.length} file(s):\n`);

  for (const file of files) {
    console.log(`Deleting: ${file.name}`);
    console.log(`  Type: ${file.mimeType}`);
    try {
      await drive.files.delete({
        fileId: file.id,
        supportsAllDrives: true
      });
      console.log(`  ✓ Deleted\n`);
    } catch (err) {
      console.log(`  ✗ Failed: ${err.message}\n`);
    }
  }

  console.log("Cleanup complete!");
};

cleanup().catch(console.error);
