function doGet(e) {
  return HtmlService
    .createHtmlOutputFromFile('index.html')
    .setTitle("Onnline Result Sytem")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function submitDT(phoneNumber, selectedSheetName, dob) {
  var ss = SpreadsheetApp.openById("1rxg5GTNX0wFkQtU3NjzkJJysBB3d7cnrCI");
  var sheet = ss.getSheetByName(selectedSheetName); 
  var dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
  var dataValues = dataRange.getValues();
  
  var ss1 = SpreadsheetApp.openById("1KGqFqnHC9t6chMZA4rWlGcRaJ8LZixv9dV");
  var sheet1 = ss1.getSheetByName("Sheet2"); 
  var dataRange1 = sheet1.getRange(2, 1, sheet1.getLastRow() - 1, sheet1.getLastColumn());
  var dataValues1 = dataRange1.getValues();
  
  var flag = 1;
  var fees = "";
  var name = "";
  var studentData = {};
  var rol = "";

  // Fetch the image URL based on phone number from Sheet2
  for (var i = 0; i < dataValues1.length; i++) {
    var storedPhoneNumber = dataValues1[i][1]; // Assuming phone numbers are in column B of Sheet2
    if (storedPhoneNumber == phoneNumber) {
      rol = dataValues1[i][0]; // Assuming image URLs are in column A of Sheet2
      break;
    }
  }
  
  for (var i = 0; i < dataValues.length; i++) {
    var storedPhoneNumber = dataValues[i][1]; // Assuming phone numbers are in column B
    var fee = dataValues[i][2];
    var search_name = dataValues[i][3];
    var dof = dataValues[i][4];
    
    if (storedPhoneNumber == phoneNumber) {
      if (dof == dob) {
        fees = fee;
        name = search_name;
        flag = 0;
        studentData = {
          roll: dataValues[i][0], // Assuming Roll Number is in column A
          rolll: dataValues[i][1],
          studentName: dataValues[i][3], // Assuming student name is in column D
          fatherName: dataValues[i][4], // Assuming father name is in column E
          imageURL: rol // Add image URL to student data
        };

        // Retrieve subject names dynamically from F10 to L10
        var subjectNamesRange = sheet.getRange("F10:M10");
        var subjectNames = subjectNamesRange.getValues()[0]; // Assuming subject names are in a single row
        
        // Loop through each subject and store obtained marks in studentData
        for (var j = 0; j < subjectNames.length; j++) {
          var subjectLabel = subjectNames[j];
          var obtainedMarks = dataValues[i][5 + j]; // Assuming obtained marks start from column F (index 5)
          studentData[subjectLabel] = obtainedMarks;
        }
        
        // Calculate total marks and total obtained marks
        var totalMaxMarks = 0;
        var totalObtainedMarks = 0;
        
        for (var k = 0; k < subjectNames.length; k++) {
          var subjectLabel = subjectNames[k];
          var maxMarks = sheet.getRange(11, 6 + k).getValue(); // Assuming max marks are in rows 11 (F11 to O11)
          var obtainedMarks = studentData[subjectLabel];
          
          totalMaxMarks += maxMarks;
          totalObtainedMarks += obtainedMarks;
        }
        
        // Calculate percentage
        var percentage = ((totalObtainedMarks / totalMaxMarks) * 100).toFixed(2);
        studentData.totalMaxMarks = totalMaxMarks;
        studentData.totalObtainedMarks = totalObtainedMarks;
        studentData.percentage = percentage;
        
        // Get grade from column P
        studentData.grade = dataValues[i][15]; // Assuming grade is in column P (index 15)
        
        // Get remarks from column Q
        studentData.remarks = dataValues[i][16]; // Assuming remarks are in column Q (index 16)
      }
    }
  }
  
  if (flag == 1) {
    return "Invalid, Please check your register number and date of birth !.";
  } else {
    if (fees == "Remove Roll Number ") {
      return "Please, enter your register number in the above field";
    }
    if (fees != "") {
      return "Dear " + name + ",We would like to inform you that your academic fees have not been fully paid. Your result can only be published upon the completion of your fees. Currently, your outstanding balance for academic fees is ₹" + fees + " .Kindly contact the management of the school at your earliest convenience to complete the outstanding fee.";
    }
    var data = "<table class='data-box'><tbody>";
        /*<center><img src="https://i.ibb.co/bHn78CG/svv-removebg-preview.png" alt="Logo" width="180" height="130">
      <p style="color:#00008b;font-family:cambria;font-size:20px;">online result system</p>
      <p style="color:#228B22;font-family:cambria;font-size:24px;">Exam Result</p></center>
      <br>
      <br>
      <br>*/
    data += `
      <tr>
        <td><strong>Name:</strong></td>
        <td>${studentData.studentName}</td>
        <td rowspan="3" style="text-align: right; padding-right: 100px;">
        <img src="${studentData.roll}" alt="Student Image" width="100" height="100">
      </tr>
      <tr>
        <td><strong>Date Of Birth:</strong></td>
        <td>${studentData.fatherName}</td>
      </tr>
      <tr>
        <td><strong>Register Number:</strong></td>
        <td>${studentData.rolll}</td>
      </tr>
      <tr>
        <td><strong>Serial</strong></td>
        <td><strong>Subjects</strong></td>
        <td><strong>Max Marks</strong></td>
        <td><strong>Obtained Marks</strong></td>
        <td><strong>Status</strong></td>
      </tr>
    `;

    // Define a function to check pass/fail status and apply color
    function getStatusCell(marks, maxMarks) {
      var percentage = (marks / maxMarks) * 100;
      var status = percentage >= 40 ? "Pass" : "Fail";
      var color = percentage >= 40 ? "green" : "red";
      if (maxMarks == 0 && marks == 0) {
        status = "-------";
        color = "yellow";
      }
      if (percentage == 0) {
        status = "Absent";
        color = "red";
      }
      return `<strong style="color: ${color};">${status}</strong>`;
    }

    // Loop through each subject and calculate status
    for (var j = 0; j < subjectNames.length; j++) {
      var subjectLabel = subjectNames[j];
      var maxMarks = sheet.getRange(11, 6 + j).getValue(); // Assuming max marks are in rows 11 (F11 to O11)
      var obtainedMarks = studentData[subjectLabel];
      
      data += `
        <tr>
          <td>${j + 1}</td>
          <td><strong>${subjectLabel}</strong></td>
          <td>${maxMarks}</td>
          <td>${obtainedMarks}</td>
          <td>${getStatusCell(obtainedMarks, maxMarks)}</td>
        </tr>
      `;
    }

    data += `
      <tr>
        <td><strong>Total Marks:</strong></td>
        <td></td>
        <td>${studentData.totalMaxMarks}</td>
        <td>${studentData.totalObtainedMarks}</td>
      </tr>
      <tr>
        <td>Grade:</td>
        <td><span style="color: green;">${studentData.grade}</span></td>
      </tr>
      <tr>
        <td>Percentage:</td>
        <td>${studentData.percentage}%</td>
      </tr>
      <tr>
        <td>Remarks:</td>
        <td>${studentData.remarks}</td>
      </tr>
    `;

    data += "</tbody></table>";

    // Add the image below the marks table
    data += `
      <br>
      <center>
        <img src="${studentData.imageURL}" alt="Student Progress Graph" width=600,height=400>
      </center>
      <br>
      <hr>
    `;
    
    return data;
  }
}
