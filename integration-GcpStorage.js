const { Storage } = require('@google-cloud/storage');
const storage = new Storage();
const bucketName = 'bot-uploads-bucket'; // Your bucket name from the design

async function uploadFile(filePath, destinationFileName) {
    await storage.bucket(bucketName).upload(filePath, {
        destination: destinationFileName,
    });
    console.log(`${filePath} uploaded to ${bucketName}/${destinationFileName}`);
}

async function downloadFile(sourceFileName, destinationPath) {
    await storage.bucket(bucketName).file(sourceFileName).download({
        destination: destinationPath,
    });
    console.log(`gs://${bucketName}/${sourceFileName} downloaded to ${destinationPath}`);
}

// In your bot logic, instead of:
// fs.writeFileSync('./uploads/my_file.pdf', data);
// You would use:
// await uploadFile('/tmp/my_temp_file.pdf', 'my_file.pdf');
// (Note: Cloud Run instances have a /tmp directory for temporary local storage)

// And instead of:
// fs.readFileSync('./uploads/my_file.pdf');
// You would use:
// await downloadFile('my_file.pdf', '/tmp/downloaded_file.pdf');

module.exports = { uploadFile, downloadFile };