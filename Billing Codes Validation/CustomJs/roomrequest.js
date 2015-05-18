function customBillingReferenceValidation(billingRef) {

	billingRef = String(billingRef);
	billingRef = billingRef.replace(/-/g, '');
	billingRef = billingRef.replace(/ /g, '');

	if (billingRef.length < 33) {
		return false;
	}

	var validation;
     

	$.ajax({type:'POST',
		url:'../BillingCodesValidationWebService/Service.asmx/customBillingReferenceValidation', 
		data: {billingCode : billingRef}, 
		dataType: 'text',
		async: false,
		success: function(data) {
			if (data.indexOf("true") !== -1) {
				validation = true;
			} else {
				alert("The billing code you entered is invalid");
				validation = false;
			}
		}
	});

	return validation;

}
