$(document).ready(function(){
	$('#Basic_category').change(function(){
		var optionSelected = $("option:selected",this).val();
		if(optionSelected === "8"){
			$('#toggleText').show();
		}
		else
		{
			$('#toggleText').hide();
		}
	});
});