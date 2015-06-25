(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();

            $('#get-data-from-selection').click(getDataFromSelection);
            $('#write-range').click(writeRange);
        });
    };

    // Reads data from current document selection and displays a notification
    function getDataFromSelection() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    app.showNotification('The selected text is:', '"' + result.value + '"');
                } else {
                    app.showNotification('Error:', result.error.message);
                }
            }
        );
    }

    function writeRange() {
        var rangeAddress = "A1:A1";

        var ctx = new Excel.ExcelClientContext();
        var range = ctx.workbook.worksheets.getActiveWorksheet().getRange(rangeAddress);

        range.getCell(0, 0).values = "Hello World!";

        ctx.executeAsync().then(function () {
            app.showNotification("Write to Range"+rangeAddress+"is Successful!");
        }, function (error) {
            app.showNotification("Error", JSON.stringify(error));
        });
    }


})();


$(document).ready(function(e) {
    $('#1,#2,#3,#4,#5,#6,#7,#8,#9,#0').click(function(){
		var v = $(this).val();
		$('#answer').val($('#answer').val() + v);	
	});
	$('#C').click(function(){
		$('#answer').val('');
		$('#operation').val('');
		$('#operation').removeClass('activeAnswer');
		$('#equals').attr('onclick','');
	});
	$('#plus').click(function(e) { 
	
		if($('#answer').val() == ''){
			return false;
			$('#equals').attr('onclick','');
		}
		else if ( $('#operation').attr('class') == 'activeAnswer') {
			$('#operation').val( $('#operation').val() + $('#plus').val() );
			$('#answer').val('');
			$('#equals').attr('onclick','');
		}
		else{
			$('#operation').val( $('#operation').val() + $('#answer').val() + $('#plus').val() );
			$('#answer').val('');
			$('#equals').attr('onclick','');
		}
    });
	$('#subtract').click(function(e) { 
	
		if($('#answer').val() == ''){
			return false;	
			$('#equals').attr('onclick','');
		}
		else if ( $('#operation').attr('class') == 'activeAnswer') {
			$('#operation').val( $('#operation').val() + $('#subtract').val() );
			$('#answer').val('');
			$('#equals').attr('onclick','');
		}
		else{
			$('#operation').val( $('#operation').val() + $('#answer').val() + $('#subtract').val() );
			$('#answer').val('');
			$('#equals').attr('onclick','');
		}
    });
	$('#divide').click(function(e) { 
	
		if($('#answer').val() == ''){
			return false;	
			$('#equals').attr('onclick','');
		}
		else if ( $('#operation').attr('class') == 'activeAnswer') {
			$('#operation').val( $('#operation').val() + $('#divide').val() );
			$('#answer').val('');
			$('#equals').attr('onclick','');
		}
		else{
			$('#operation').val( $('#operation').val() + $('#answer').val() + $('#divide').val() );
			$('#answer').val('');
			$('#equals').attr('onclick','');
		}
    });
	$('#product').click(function(e) { 
	
		if($('#answer').val() == ''){
			return false;	
			$('#equals').attr('onclick','');
		}
		else if ( $('#operation').attr('class') == 'activeAnswer') {
			$('#operation').val( $('#operation').val() + $('#product').val() );
			$('#answer').val('');
			$('#equals').attr('onclick','');
		}
		else{
			$('#operation').val( $('#operation').val() + $('#answer').val() + $('#product').val() );
			$('#answer').val('');
			$('#equals').attr('onclick','');
		}
    });	
	$('#equals').click(function(){
		
		if($('#equals').attr('onclick') != 'return false'){
		
			var a = $('#answer').val();
			var b = $('#operation').val();
			var c = b.concat(a);
			$('#answer').val(eval(c));
			$('#operation').val(eval(c));
			$('#operation').addClass('activeAnswer');
			$('#equals').attr('onclick','return false');
		
		}
	});
});