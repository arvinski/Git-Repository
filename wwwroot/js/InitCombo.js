class OpInitCombo {

	static initReasonCode() {
		jQuery('.js-change-reasoncode').change(function () {
			var rc = $("#ReasonCode").val();
			console.log(rc);
			switch (rc) {
				case "IPC":
				case "IPR":
				case "NCC":
				case "PNC":
				case "SPR":

					$('#TransType').select2({
						disabled: false
					});

					document.getElementById("Destination").disabled = false;

					$('#ATMTrans').select2({
						disabled: false
					});

					$('#ATMUsed').select2({
						disabled: false
					});

					$('#Location').select2({
						disabled: false
					});

					$('#PaymentOf').select2({
						disabled: false
					});

					$('#TerminalUsed').select2({
						disabled: false
					});

					document.getElementById("BillerName").disabled = false;
					document.getElementById("Merchant").disabled = false;
					document.getElementById("cpPresent").disabled = false;
					document.getElementById("cpNotPresent").disabled = false;

					$('#BancnetUsed').select2({
						disabled: false
					});

					document.getElementById("Online").disabled = false;
					document.getElementById("Website").disabled = false;

					document.getElementById("Amount").disabled = false;

					$('#Currency').select2({
						disabled: false
					});

					document.getElementById("DateTrans").disabled = false;

					document.getElementById("inte").disabled = false;
					document.getElementById("inta").disabled = false;

					$('#PLStatus').select2({
						disabled: true
					});

					document.getElementById("RemitFrom").disabled = true;
					document.getElementById("RemitConcern").disabled = true;

					break;

				case "DGI":
				case "DLR":
				case "DRR":
				case "PAI":
				case "PAS":
				case "PGI":
					$('#TransType').select2({
						disabled: true
					});

					document.getElementById("Destination").disabled = true;

					$('#ATMTrans').select2({
						disabled: true
					});

					$('#ATMUsed').select2({
						disabled: true
					});

					$('#Location').select2({
						disabled: true
					});

					$('#PaymentOf').select2({
						disabled: true
					});

					$('#TerminalUsed').select2({
						disabled: true
					});

					document.getElementById("BillerName").disabled = true;
					document.getElementById("Merchant").disabled = true;
					document.getElementById("cpPresent").disabled = true;
					document.getElementById("cpNotPresent").disabled = true;

					$('#BancnetUsed').select2({
						disabled: true
					});

					document.getElementById("Online").disabled = true;
					document.getElementById("Website").disabled = true;

					document.getElementById("Amount").disabled = true;

					$('#Currency').select2({
						disabled: true
					});

					document.getElementById("DateTrans").disabled = true;

					document.getElementById("inte").disabled = true;
					document.getElementById("inta").disabled = true;

					$('#PLStatus').select2({
						disabled: false
					});

					document.getElementById("RemitFrom").disabled = true;
					document.getElementById("RemitConcern").disabled = true;

					break;

				case "FCS":
				case "BRC":
				case "FCP":

					$('#TransType').select2({
						disabled: false
					});

					document.getElementById("Destination").disabled = true;

					$('#ATMTrans').select2({
						disabled: true
					});

					$('#ATMUsed').select2({
						disabled: true
					});

					$('#Location').select2({
						disabled: true
					});

					$('#PaymentOf').select2({
						disabled: true
					});

					$('#TerminalUsed').select2({
						disabled: true
					});

					document.getElementById("BillerName").disabled = true;
					document.getElementById("Merchant").disabled = true;
					document.getElementById("cpPresent").disabled = true;

					document.getElementById("cpNotPresent").disabled = true;


					$('#BancnetUsed').select2({
						disabled: true
					});

					document.getElementById("Online").disabled = true;
					document.getElementById("Website").disabled = true;

					document.getElementById("Amount").disabled = true;

					$('#Currency').select2({
						disabled: true
					});

					document.getElementById("DateTrans").disabled = true;

					document.getElementById("inte").disabled = true;

					document.getElementById("inta").disabled = true;

					document.getElementById("RemitFrom").disabled = true;
					document.getElementById("RemitConcern").disabled = true;

					$('#PLStatus').select2({
						disabled: true
					});


					break;

				default:

					$('#TransType').select2({
						disabled: true
					});

					document.getElementById("Destination").disabled = true;

					$('#ATMTrans').select2({
						disabled: true
					});

					$('#Location').select2({
						disabled: true
					});

					$('#ATMUsed').select2({
						disabled: true
					});

					$('#PaymentOf').select2({
						disabled: true
					});

					$('#TerminalUsed').select2({
						disabled: true
					});

					document.getElementById("BillerName").disabled = true;
					document.getElementById("Merchant").disabled = true;
					document.getElementById("cpPresent").disabled = true;
					document.getElementById("cpNotPresent").disabled = true;

					$('#BancnetUsed').select2({
						disabled: true
					});

					document.getElementById("Online").disabled = true;
					document.getElementById("Website").disabled = true;
					document.getElementById("Website").disabled = true;

					document.getElementById("Amount").disabled = true;
					$('#Currency').select2({
						disabled: true
					});

					document.getElementById("DateTrans").disabled = true;

					document.getElementById("inte").disabled = true;
					document.getElementById("inta").disabled = true;

					document.getElementById("RemitFrom").disabled = true;
					document.getElementById("RemitConcern").disabled = true;

					$('#PLStatus').select2({
						disabled: true
					});

					break;

			}
		});

	}

	static initStatus() {
		jQuery('.js-change-status').change(function () {

			var st = $("#Status").val();
			/*$("#SubStatus").empty();*/

			switch (st) {
				case "Endorsed":
					//$('#SubStatus').select2({
					//	disabled: false
					//});

					$('#EndoredTo').select2({
						disabled: false
					});

					$('#Remarks').select2({
						disabled: false
					});

					document.getElementById("Resolved").disabled = false;

					$('#ReferedBy').select2({
						disabled: false
					});

					//document.getElementById("ContactPerson").disabled = false;

					break;

				case "Commitment":

					//$('#SubStatus').select2({
					//	disabled: true
					//});

					$('#EndoredTo').select2({
						disabled: true
					});

					//$('#Remarks').select2({
					//	disabled: true
					//});

					document.getElementById("Resolved").disabled = false;

					$('#ReferedBy').select2({
						disabled: false
					});

					//document.getElementById("ContactPerson").disabled = true;

					break;

				default:

					//$('#SubStatus').select2({
					//	disabled: true
					//});

					$('#EndoredTo').select2({
						disabled: true
					});

					//$('#Remarks').select2({
					//	disabled: true
					//});

					//document.getElementById("Resolved").disabled = true;

					$('#ReferedBy').select2({
						disabled: true
					});

				//document.getElementById("ContactPerson").disabled = true;
			}



		});
	}

	static initTagging() {
		jQuery('.js-change-tagging').change(function () {
			var ch = $("#Tagging").val();
			switch (ch) {
				case "SIMPLE":

					var simple_date = new Date();

					simple_date.setDate(simple_date.getDate() + 7)

					$('#Resolved').datepicker({
						format: 'dd-M-yyyy',
						autoclose: true
					});

					$('#Resolved').datepicker('setDate', simple_date);

					var rc = $("#ReasonCode").val();

					if (rc == "FSI") {

						$('#EndoredTo').select2({
							disabled: true
						});

						$("#EndoredTo").val('').trigger('change');

						$('#EndorsedFrom').select2({
							disabled: true
						});

						$("#EndorsedFrom").val('').trigger('change');

						$('#Remarks').select2({
							disabled: true
						});

						$("#Remarks").val('').trigger('change');

						$('#ReferedBy').select2({
							disabled: true
						});

						$("#ReferedBy").val('').trigger('change');

						document.getElementById("CallTime").value = "";
					}


					break;

				case "COMPLEX":

					var complx_date = new Date();

					complx_date.setDate(complx_date.getDate() + 45)

					$('#Resolved').datepicker({
						format: 'dd-M-yyyy',
						autoclose: true
					});

					$('#Resolved').datepicker('setDate', complx_date);

					var today = new Date();
					var hour = today.getHours();
					var minute = today.getMinutes();

					var rc = $("#ReasonCode").val();

					if (rc == "FSI") {

						$('#EndoredTo').select2({
							disabled: false,
							allowClear: true
						});

						$('#EndorsedFrom').select2({
							disabled: false,
							allowClear: true
						});


						$('#Remarks').select2({
							disabled: false,
							allowClear: true

						});

						$('#ReferedBy').select2({
							disabled: false,
							allowClear: true
						});


						document.getElementById("CallTime").value = String(hour).padStart(2, "0") + ":" + String(minute).padStart(2, "0");

					}

					break;

				case "COMPLEX - EMV":

					var cmplx_emv_date = new Date();

					cmplx_emv_date.setDate(cmplx_emv_date.getDate() + 10)

					$('#Resolved').datepicker({
						format: 'dd-M-yyyy',
						autoclose: true
					});

					$('#Resolved').datepicker('setDate', cmplx_emv_date);

					var today = new Date();
					var hour = today.getHours();
					var minute = today.getMinutes();

					var rc = $("#ReasonCode").val();

					if (rc == "FSI") {

						$('#EndoredTo').select2({
							disabled: false,
							allowClear: true
						});

						$('#EndorsedFrom').select2({
							disabled: false,
							allowClear: true
						});


						$('#Remarks').select2({
							disabled: false,
							allowClear: true

						});

						$('#ReferedBy').select2({
							disabled: false,
							allowClear: true
						});


						document.getElementById("CallTime").value = String(hour).padStart(2, "0") + ":" + String(minute).padStart(2, "0");

					}

					break;

				case "CCOMPLEX - DATA PRIVACY":

					var cmplx_dta_date = new Date();

					cmplx_dta_date.setDate(cmplx_dta_date.getDate() + 26)

					$('#Resolved').datepicker({
						format: 'dd-M-yyyy',
						autoclose: true
					});

					$('#Resolved').datepicker('setDate', cmplx_dta_date);

					var today = new Date();
					var hour = today.getHours();
					var minute = today.getMinutes();

					var rc = $("#ReasonCode").val();

					if (rc == "FSI") {

						$('#EndoredTo').select2({
							disabled: false,
							allowClear: true
						});

						$('#EndorsedFrom').select2({
							disabled: false,
							allowClear: true
						});


						$('#Remarks').select2({
							disabled: false,
							allowClear: true

						});

						$('#ReferedBy').select2({
							disabled: false,
							allowClear: true
						});


						document.getElementById("CallTime").value = String(hour).padStart(2, "0") + ":" + String(minute).padStart(2, "0");

					}

					break;

			}

		});
	}

	static initChannel() {
		jQuery('.js-change-transtype').change(function () {
			var ch = $("#TransType").val();
			switch (ch) {
				case "ATM":

					document.getElementById("Destination").disabled = true;

					$('#ATMTrans').select2({
						disabled: false
					});

					$('#Location').select2({
						disabled: false
					});

					$('#ATMUsed').select2({
						disabled: false
					});

					$('#PaymentOf').select2({
						disabled: true
					});

					$('#TerminalUsed').select2({
						disabled: true
					});

					document.getElementById("BillerName").disabled = true;
					document.getElementById("Merchant").disabled = true;
					document.getElementById("cpPresent").disabled = true;
					
					document.getElementById("cpNotPresent").disabled = true;
					

					$('#BancnetUsed').select2({
						disabled: true
					});

					document.getElementById("Online").disabled = true;
					document.getElementById("Website").disabled = true;

					document.getElementById("Amount").disabled = false;
					$('#Currency').select2({
						disabled: false
					});

					document.getElementById("DateTrans").disabled = false;

					document.getElementById("inte").disabled = true;
					
					document.getElementById("inta").disabled = true;
					

					document.getElementById("RemitFrom").disabled = true;
					document.getElementById("RemitConcern").disabled = true;

					$('#PLStatus').select2({
						disabled: false
					});


					break;
				case "BILLS PAYMENT":
					document.getElementById("Destination").disabled = true;

					$('#ATMTrans').select2({
						disabled: true
					});

					$('#Location').select2({
						disabled: true
					});

					$('#ATMUsed').select2({
						disabled: true
					});

					$('#PaymentOf').select2({
						disabled: false
					});

					$('#TerminalUsed').select2({
						disabled: true
					});

					document.getElementById("BillerName").disabled = false;
					document.getElementById("Merchant").disabled = false;
					document.getElementById("cpPresent").disabled = true;
					
					document.getElementById("cpNotPresent").disabled = true;
					

					$('#BancnetUsed').select2({
						disabled: false
					});

					document.getElementById("Online").disabled = false;
					document.getElementById("Website").disabled = false;

					document.getElementById("Amount").disabled = false;
					$('#Currency').select2({
						disabled: false
					});

					document.getElementById("DateTrans").disabled = false;

					document.getElementById("inte").disabled = true;
					
					document.getElementById("inta").disabled = true;
					

					document.getElementById("RemitFrom").disabled = false;
					document.getElementById("RemitConcern").disabled = false;

					$('#PLStatus').select2({
						disabled: false
					});

					break;

				case "IBFT":

					document.getElementById("Destination").disabled = false;

					$('#ATMTrans').select2({
						disabled: true
					});

					$('#Location').select2({
						disabled: true
					});

					$('#ATMUsed').select2({
						disabled: true
					});

					$('#PaymentOf').select2({
						disabled: true
					});

					$('#TerminalUsed').select2({
						disabled: true
					});

					document.getElementById("BillerName").disabled = true;
					document.getElementById("Merchant").disabled = true;
					document.getElementById("cpPresent").disabled = true;
					
					document.getElementById("cpNotPresent").disabled = true;
					

					$('#BancnetUsed').select2({
						disabled: false
					});

					document.getElementById("Online").disabled = false;
					document.getElementById("Website").disabled = false;

					document.getElementById("Amount").disabled = false;
					$('#Currency').select2({
						disabled: false
					});

					document.getElementById("DateTrans").disabled = false;

					document.getElementById("inte").disabled = false;
					
					document.getElementById("inta").disabled = false;
					

					document.getElementById("RemitFrom").disabled = false;
					document.getElementById("RemitConcern").disabled = false;

					$('#PLStatus').select2({
						disabled: false
					});

					break;

				case "POS":

					document.getElementById("Destination").disabled = true;

					$('#ATMTrans').select2({
						disabled: true
					});

					$('#Location').select2({
						disabled: true
					});

					$('#ATMUsed').select2({
						disabled: true
					});

					$('#PaymentOf').select2({
						disabled: true
					});

					$('#TerminalUsed').select2({
						disabled: false
					});

					document.getElementById("BillerName").disabled = false;
					document.getElementById("Merchant").disabled = false;
					document.getElementById("cpPresent").disabled = false;
					
					document.getElementById("cpNotPresent").disabled = false;
					

					$('#BancnetUsed').select2({
						disabled: true
					});

					document.getElementById("Online").disabled = true;
					document.getElementById("Website").disabled = true;

					document.getElementById("Amount").disabled = false;
					$('#Currency').select2({
						disabled: false
					});

					document.getElementById("DateTrans").disabled = false;

					document.getElementById("inte").disabled = true;
					
					document.getElementById("inta").disabled = true;
					

					document.getElementById("RemitFrom").disabled = false;
					document.getElementById("RemitConcern").disabled = false;

					$('#PLStatus').select2({
						disabled: false
					});

					break;

				case "REMITTANCE":

					document.getElementById("Destination").disabled = true;

					$('#ATMTrans').select2({
						disabled: true
					});

					$('#Location').select2({
						disabled: true
					});

					$('#ATMUsed').select2({
						disabled: true
					});

					$('#PaymentOf').select2({
						disabled: true
					});

					$('#TerminalUsed').select2({
						disabled: true
					});

					document.getElementById("BillerName").disabled = true;
					document.getElementById("Merchant").disabled = true;
					document.getElementById("cpPresent").disabled = true;
					document.getElementById("cpNotPresent").disabled = true;

					$('#BancnetUsed').select2({
						disabled: true
					});

					document.getElementById("Online").disabled = true;
					document.getElementById("Website").disabled = true;

					document.getElementById("Amount").disabled = false;
					$('#Currency').select2({
						disabled: false
					});

					document.getElementById("DateTrans").disabled = false;

					document.getElementById("inte").disabled = true;
					document.getElementById("inta").disabled = true;

					document.getElementById("RemitFrom").disabled = false;
					document.getElementById("RemitConcern").disabled = false;

					$('#PLStatus').select2({
						disabled: false
					});

					break;

			}

		});
	}

	static initATMUsed() {
		jQuery('.js-change-atmused').change(function () {
			/*$("#Location").empty();*/

			//$.ajax({
			//	type: 'GET',
			//	/*url: '@Url.Action("GetATMLocation", "Activity")',*/
			//	url: '/Activity/GetATMLocation',
			//	dataType: 'json',
			//	data: { ATMUsed: $("#ATMUsed").val() },
			//	success: function (location) {

			//		/*$("#Location").append('<option value="' + -1 + '">' + "Choose ATM Branch Location" + '</option>');*/
			//		$.each(location, function (i, row) {
			//			console.log(i, row);
			//			$("#Location").append('<option value="' + row.value + '">' + row.text + '</option>');
			//		});
			//	},
			//	error: function (ex) {
			//		alert('Failed to retrieve Location.' + ex);
			//	}
			//});


			var au = $("#ATMUsed").val();

			switch (au) {
				case "325":
					$('#Location').next(".select2-container").show();

					document.getElementById('location2').style.display = 'none';
					break;
				default:
					$('#Location').next(".select2-container").hide();

					document.getElementById('location2').style.display = 'block';
					break;
			}


		});
	}

	static initEndored() {
		jQuery('.js-change-endoredto').change(function () {
			var et = $("#EndoredTo").val();
			if (et == "Others") {
				$('#EndorsedFrom').select2({
					disabled: false
				});
			}
			else {
				$('#EndorsedFrom').select2({
					disabled: true
				});

			}
		});
	}

    static init() {
		this.initReasonCode();
		this.initStatus();
		this.initTagging();
		this.initChannel();
		this.initATMUsed();
		this.initEndored();
    }
}

// Initialize when page loads
jQuery(() => { OpInitCombo.init(); });