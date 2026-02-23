function doPost(e) {
  try {
    Logger.log('=== FORM SUBMISSION STARTED ===');
    
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = JSON.parse(e.postData.contents);
    
    Logger.log('Email: ' + data.email);
    Logger.log('Send Copy: ' + data.sendCopy);
    
    if (sheet.getLastRow() === 0) {
      var headers = ['Email', 'Your Name', 'How Informative (1-5)', 
                     'Impact Coverage (1-5)', 'IT Projects Interested', 'Mobile Apps Interested', 
                     'Overall Design Rating (1-5)', 'Additional Feedback', 'Send Copy', 'Timestamp'];
      sheet.appendRow(headers);
      var headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#4285f4');
      headerRange.setFontColor('#ffffff');
    }
  
    var timestamp = new Date();
    var rowData = [
      data.email || '',
      data.name || '',
      data.informative || '',
      data.impact || '',
      Array.isArray(data.projects) ? data.projects.join('\n') : '',
      Array.isArray(data.apps) ? data.apps.join('\n') : '',
      data.design || '',
      data.feedback || '',
      data.sendCopy ? 'Yes' : 'No',
      timestamp
    ];
    
    sheet.appendRow(rowData);
    Logger.log('Data saved to sheet');
  
    var emailResult = 'not_requested';
    if (data.sendCopy && data.email) {
      Logger.log('Sending email to: ' + data.email);
      try {
        sendEmailToUser(data.email, data);
        emailResult = 'sent';
        Logger.log('Email sent successfully!');
      } catch (emailError) {
        Logger.log('Email error: ' + emailError.toString());
        emailResult = 'failed';
      }
    }
    
    return ContentService
      .createTextOutput(JSON.stringify({
        'result': 'success',
        'message': 'Form submitted successfully!',
        'emailStatus': emailResult
      }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    Logger.log('ERROR: ' + error.toString());
    return ContentService
      .createTextOutput(JSON.stringify({
        'result': 'error',
        'message': error.toString()
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function sendEmailToUser(email, data) {
  var subject = 'Your response on Chhatrapati Sambhajinagar CTB';
  
  var plainBody = 'Maha Infotech Pvt Ltd Chhatrapati Sambhajinagar CTB Feedback\n\n';
  plainBody += 'Email\n' + (data.email || '') + '\n\n';
  plainBody += 'Your Name\n' + (data.name || '') + '\n\n';
  plainBody += 'How informative did you find the CTB?\n' + (data.informative || '') + '\n\n';
  plainBody += 'Did you think the CTB covered the impact of the digital projects on Chhatrapati Sambhajinagar?\n' + (data.impact || '') + '\n\n';
  
  if (Array.isArray(data.projects) && data.projects.length > 0) {
    plainBody += 'Which of the IT projects covered in the CTB did you find interesting and would like more information on?\n';
    plainBody += data.projects.join(', ') + '\n\n';
  } else {
    plainBody += 'Which of the IT projects covered in the CTB did you find interesting and would like more information on?\n\n';
  }
  
  if (Array.isArray(data.apps) && data.apps.length > 0) {
    plainBody += 'Which of the Mobile apps included in the CTB did you find interesting and would like more information on?\n';
    plainBody += data.apps.join(', ') + '\n\n';
  } else {
    plainBody += 'Which of the Mobile apps included in the CTB did you find interesting and would like more information on?\n\n';
  }
  
  plainBody += 'How happy are you with the overall design and layout of the CTB?\n' + (data.design || '') + '\n\n';
  
  if (data.feedback && data.feedback.trim() !== '') {
    plainBody += 'Please provide any other feedback that you may like to share about the CTB or any of the projects at Chhatrapati Sambhajinagar!\n' + data.feedback + '\n\n';
  } else {
    plainBody += 'Please provide any other feedback that you may like to share about the CTB or any of the projects at Chhatrapati Sambhajinagar!\n\n';
  }
  
  plainBody += 'This form was created inside of MIPL.';
  
  // Get header image from Google Drive
  // Get header image from Google Drive
  // Using direct URL approach for better email compatibility
  var headerFileId = '1Dna7c6_I1T30MnzikSkDscFBR56LIYFd';
  var headerImageUrl = 'https://drive.google.com/uc?export=view&id=' + headerFileId;

  var htmlBody = '<!DOCTYPE html>';
  htmlBody += '<html><head><meta charset="UTF-8"></head>';
  htmlBody += '<body style="margin: 0; padding: 0; font-family: Roboto, Arial, sans-serif; background-color: #f5f5f5;">';
  htmlBody += '<div style="max-width: 700px; margin: 0 auto; background-color: #ffffff;">';
  
  // Header image with same styling as webpage
  htmlBody += '<div style="width: 100%; height: 200px; overflow: hidden; margin: 0; padding: 0; border-radius: 12px 12px 0 0;">';
  htmlBody += '<img src="' + headerImageUrl + '" alt="Chhatrapati Sambhajinagar" style="width: 100%; height: 100%; display: block; object-fit: cover; object-position: center 20%;">';
  htmlBody += '</div>';
  
  // Title section
  htmlBody += '<div style="padding: 24px 24px 20px 24px; border-bottom: 1px solid #dadce0;">';
  htmlBody += '<h2 style="margin: 0; font-size: 24px; font-weight: 400; color: #202124; line-height: 1.3;">Feedback on the Digital Coffee Table Book of Chhatrapati Sambhajinagar</h2>';
  htmlBody += '</div>';
  
  // Email field
  htmlBody += '<div style="padding: 24px; border-bottom: 1px solid #dadce0;">';
  htmlBody += '<div style="margin-bottom: 4px; font-size: 14px; color: #70757a;">Email</div>';
  htmlBody += '<div style="font-size: 15px; color: #202124;">' + (data.email || '') + '</div>';
  htmlBody += '</div>';
 
  // Name field
  htmlBody += '<div style="padding: 24px; border-bottom: 1px solid #dadce0;">';
  htmlBody += '<div style="margin-bottom: 4px; font-size: 14px; color: #70757a;">Your Name</div>';
  htmlBody += '<div style="font-size: 15px; color: #202124;">' + (data.name || '') + '</div>';
  htmlBody += '</div>';
  
  // Question 1
  htmlBody += '<div style="padding: 24px; border-bottom: 1px solid #dadce0;">';
  htmlBody += '<div style="margin-bottom: 8px; font-size: 14px; color: #70757a;">How informative did you find the CTB?</div>';
  htmlBody += '<div style="display: inline-block; background: #1a73e8; color: white; padding: 6px 12px; border-radius: 16px; font-size: 14px; font-weight: 600;">';
  htmlBody += (data.informative || '') + ' / 5';
  htmlBody += '</div></div>';
  
  // Question 2
  htmlBody += '<div style="padding: 24px; border-bottom: 1px solid #dadce0;">';
  htmlBody += '<div style="margin-bottom: 8px; font-size: 14px; color: #70757a;">Did you think the CTB covered the impact of the digital projects on Chhatrapati Sambhajinagar?</div>';
  htmlBody += '<div style="display: inline-block; background: #1a73e8; color: white; padding: 6px 12px; border-radius: 16px; font-size: 14px; font-weight: 600;">';
  htmlBody += (data.impact || '') + ' / 5';
  htmlBody += '</div></div>';
  
  // Question 3
  htmlBody += '<div style="padding: 24px; border-bottom: 1px solid #dadce0;">';
  htmlBody += '<div style="margin-bottom: 8px; font-size: 14px; color: #70757a;">Which of the IT projects covered in the CTB did you find interesting and would like more information on? If you need more information on a specific project, please use the option "Other" and include the name of the project.</div>';
  if (Array.isArray(data.projects) && data.projects.length > 0) {
    data.projects.forEach(function(project) {
      htmlBody += '<div style="font-size: 15px; color: #202124; margin-bottom: 4px;">• ' + project + '</div>';
    });
  } else {
    htmlBody += '<div style="font-size: 15px; color: #202124;"></div>';
  }
  htmlBody += '</div>';
  
  // Question 4
  htmlBody += '<div style="padding: 24px; border-bottom: 1px solid #dadce0;">';
  htmlBody += '<div style="margin-bottom: 8px; font-size: 14px; color: #70757a;">Which of the Mobile apps included in the CTB did you find interesting and would like more information on?</div>';
  if (Array.isArray(data.apps) && data.apps.length > 0) {
    data.apps.forEach(function(app) {
      htmlBody += '<div style="font-size: 15px; color: #202124; margin-bottom: 4px;">• ' + app + '</div>';
    });
  } else {
    htmlBody += '<div style="font-size: 15px; color: #202124;"></div>';
  }
  htmlBody += '</div>';
  
  // Question 5
  htmlBody += '<div style="padding: 24px; border-bottom: 1px solid #dadce0;">';
  htmlBody += '<div style="margin-bottom: 8px; font-size: 14px; color: #70757a;">How happy are you with the overall design and layout of the CTB?</div>';
  htmlBody += '<div style="display: inline-block; background: #1a73e8; color: white; padding: 6px 12px; border-radius: 16px; font-size: 14px; font-weight: 600;">';
  htmlBody += (data.design || '') + ' / 5';
  htmlBody += '</div></div>';
  
  // Question 6
  htmlBody += '<div style="padding: 24px; border-bottom: 1px solid #dadce0;">';
  htmlBody += '<div style="margin-bottom: 4px; font-size: 14px; color: #70757a;">Please provide any other feedback that you may like to share about the CTB or any of the projects at Chhatrapati Sambhajinagar!</div>';
  htmlBody += '<div style="font-size: 15px; color: #202124; white-space: pre-wrap;">' + (data.feedback && data.feedback.trim() !== '' ? data.feedback : '') + '</div>';
  htmlBody += '</div>';
  
  // Footer
  htmlBody += '<div style="padding: 24px; text-align: center; background-color: #f5f5f5;">';
  htmlBody += '<p style="margin: 0; color: #5f6368; font-size: 12px;">This form was created inside of MIPL.</p>';
  htmlBody += '</div>';
  
  htmlBody += '</div></body></html>';
 
  var emailOptions = {
    htmlBody: htmlBody,
    name: 'Maha Infotech Pvt Ltd Chhatrapati Sambhajinagar CTB Feedback'
  };
  
  GmailApp.sendEmail(email, subject, plainBody, emailOptions);
  
  return true;
}

// STEP 1: Run this function FIRST to authorize email sending and Drive access
function setupEmailPermissions() {
  Logger.log('Setting up email and Drive permissions...');
  
  try {
    var myEmail = Session.getEffectiveUser().getEmail();
    Logger.log('Your email: ' + myEmail);
    
    var headerFileId = '1Dna7c6_I1T30MnzikSkDscFBR56LIYFd';
    try {
      var file = DriveApp.getFileById(headerFileId);
      Logger.log('Drive access OK - File found: ' + file.getName());
    } catch (driveError) {
      Logger.log('Drive access issue: ' + driveError.toString());
    }
    
    GmailApp.sendEmail(
      myEmail,
      'Email Authorization Complete - CTB Feedback Form',
      'SUCCESS!\n\nIf you receive this email, the authorization is complete.\n\nYour feedback form will now send emails to users with their responses when they check "Send me a copy".\n\nTest the form now!'
    );
    
    Logger.log('SUCCESS! Email sent to: ' + myEmail);
    Logger.log('Authorization complete. Deploy the script and test your form.');
    return 'SUCCESS - Check your email!';
    
  } catch (error) {
    Logger.log('ERROR: ' + error.toString());
    Logger.log('Click "Review Permissions" and authorize the script.');
    return 'FAILED - Authorization needed';
  }
}

// Test function - sends a sample email with fake form data
function testEmailWithSampleData() {
  var testEmail = 'swarali.pathrikar@consultmipl.com'; // Change to your email
  
  Logger.log('Testing email with sample data...');
  Logger.log('Sending to: ' + testEmail);
  
  var sampleData = {
    email: testEmail,
    name: 'Test User',
    informative: '5',
    impact: '4',
    projects: ['Governance Projects', 'Citizen Centric Projects'],
    apps: ['Smart Nagrik', 'Smart Chhatrapati Sambhajinagar WhatsApp Chatbot'],
    design: '5',
    feedback: 'This is a test feedback message to see how the email looks.',
    sendCopy: true
  };
  
  try {
    sendEmailToUser(testEmail, sampleData);
    Logger.log('Test email sent successfully!');
    Logger.log('Check your inbox at: ' + testEmail);
    return 'SUCCESS - Check your email!';
  } catch (error) {
    Logger.log('Test failed: ' + error.toString());
    return 'FAILED: ' + error.toString();
  }
}

// Test Drive access - Run this to check if the header image can be accessed
function testDriveAccess() {
  var headerFileId = '1Dna7c6_I1T30MnzikSkDscFBR56LIYFd';
  
  Logger.log('Testing Drive access...');
  Logger.log('File ID: ' + headerFileId);
  
  try {
    var file = DriveApp.getFileById(headerFileId);
    Logger.log('SUCCESS! File found: ' + file.getName());
    Logger.log('File size: ' + file.getSize() + ' bytes');
    Logger.log('File type: ' + file.getMimeType());
    
    var blob = file.getBlob();
    Logger.log('Blob created successfully');
    Logger.log('Blob size: ' + blob.getBytes().length + ' bytes');
    
    return 'SUCCESS - File accessible!';
  } catch (error) {
    Logger.log('ERROR accessing file: ' + error.toString());
    Logger.log('Possible solutions:');
    Logger.log('1. Make sure you run setupEmailPermissions() first to authorize Drive access');
    Logger.log('2. Check that the file ID is correct');
    Logger.log('3. Verify the file exists in your Google Drive');
    return 'FAILED: ' + error.toString();
  }
}

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({
      'result': 'success',
      'message': 'Web app is running!'
    }))
    .setMimeType(ContentService.MimeType.JSON);
}
