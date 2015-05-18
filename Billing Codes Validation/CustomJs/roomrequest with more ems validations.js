function customBillingReferenceValidation(billingRef) {

	var validationNeeded = false;

	var dayOfTheWeek = document.getElementById('recurrenceText').innerHTML;
	if (dayOfTheWeek == "") {		
		dayOfTheWeek = document.getElementById('ctl00_pc_BookDate_box').value;
	}
	if (dayOfTheWeek.indexOf("Sun") !== -1 || dayOfTheWeek.indexOf("Sat") !== -1) {
		validationNeeded = true;	
	}

	var startHour = document.getElementById('ctl00_pc_StartTime_box').value;
	
	if ( (startHour.indexOf("10:") !== -1 && startHour.indexOf("PM") !== -1) || (startHour.indexOf("11:") !== -1 && startHour.indexOf("PM") !== -1) || (startHour.indexOf("12:") !== -1 && startHour.indexOf("AM") !== -1) || (startHour.indexOf("1:") !== -1 && startHour.indexOf("AM") !== -1) || (startHour.indexOf("2:") !== -1 && startHour.indexOf("AM") !== -1) || (startHour.indexOf("3:") !== -1 && startHour.indexOf("AM") !== -1) || (startHour.indexOf("4:") !== -1 && startHour.indexOf("AM") !== -1) || (startHour.indexOf("5:") !== -1 && startHour.indexOf("AM") !== -1) || (startHour.indexOf("6:") !== -1 && startHour.indexOf("AM") !== -1) ) {
		validationNeeded = true;	
	}

	var endHour = document.getElementById('ctl00_pc_EndTime_box').value;
	
		if ( (endHour.indexOf("10:") !== -1 && endHour.indexOf("PM") !== -1) || (endHour.indexOf("11:") !== -1 && endHour.indexOf("PM") !== -1) || (endHour.indexOf("12:") !== -1 && endHour.indexOf("AM") !== -1) || (endHour.indexOf("1:") !== -1 && endHour.indexOf("AM") !== -1) || (endHour.indexOf("2:") !== -1 && endHour.indexOf("AM") !== -1) || (endHour.indexOf("3:") !== -1 && endHour.indexOf("AM") !== -1) || (endHour.indexOf("4:") !== -1 && endHour.indexOf("AM") !== -1) || (endHour.indexOf("5:") !== -1 && endHour.indexOf("AM") !== -1) || (endHour.indexOf("6:") !== -1 && endHour.indexOf("AM") !== -1) ) {
		validationNeeded = true;	
	}
	
	var setupNotes = document.getElementById('ctl00_pc_CategoryRepeater_ctl00_ctl00_NoteBox').value;
	
	var validation;
	
	if (validationNeeded == true) {
	
		billingRef = String(billingRef);
		billingRef = billingRef.replace(/-/g, '');
		billingRef = billingRef.replace(/ /g, '');
	
		if (billingRef.length < 33) {
			alert("The billing code you entered is invalid, it should be 33 digits long");
			return false;
		}
	
		$.ajax({type:'POST',
			url:'../BillingCodesValidationWebService/Service.asmx/customBillingReferenceValidation', 
			data: {billingCode : billingRef}, 
			dataType: 'text',
			async: false,
			success: function(data) {
				if (data.indexOf("true") !== -1) {
					validationResult = true;
				} else {
					alert("The billing code you entered is invalid");
					validationResult = false;
				}
			}
		});
		
	} else {
		
		document.getElementById('ctl00_pc_ResBillingReference_box').value = "Not Required";
		validationResult = true;
		
	}
	
	return validationResult;
	

}
