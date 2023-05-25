// Google App Script

function saveEmailsAsPDF() {
    var labelName = 'Elliott/Caspari';
    var keyword = 'support';
    var startDate = new Date('August 1, 2022');
    var endDate = new Date('May 25, 2023');
    var folderName = 'Elliott Support Convo';
  
    // Get all emails matching the criteria
    var threads = GmailApp.search('label:' + labelName + ' after:' + startDate.toISOString() + ' before:' + endDate.toISOString() + ' "' + keyword + '"');
    var messages = GmailApp.getMessagesForThreads(threads);
    var emails = messages.reduce(function(a, b) { return a.concat(b); });
  
    // Create or retrieve the destination folder
    var destinationFolder = getOrCreateFolder(folderName);
  
    // Sort emails by date in ascending order
    emails.sort(function(a, b) { return a.getDate().valueOf() - b.getDate().valueOf(); });
  
    // Convert emails to PDF and save in the destination folder
    for (var i = 0; i < emails.length; i++) {
      var email = emails[i];
      var pdfName = 'Email_' + (i + 1) + '.pdf';
      var pdf = email.getAs('application/pdf');
      pdf.setName(pdfName);
      destinationFolder.createFile(pdf);
    }
  
    Logger.log('Emails saved as PDFs in the folder: ' + destinationFolder.getName());
  }
  
  function getOrCreateFolder(folderName) {
    var folder;
    var folders = DriveApp.getFoldersByName(folderName);
  
    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      folder = DriveApp.createFolder(folderName);
    }
  
    return folder;
  }
  