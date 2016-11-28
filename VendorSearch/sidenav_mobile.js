$(document).ready(function() {
	
	 $('.mobile-menu a i').click(function(){
        var el = $(this);
        el.closest('.mobile-container').find('.subpage-menu').slideToggle();
       
        if(el.hasClass('fa-caret-right')) {
        	el.removeClass('fa-caret-right');
        } else {
        	el.addClass('fa-caret-right');
        }
   });

	$('#Basic_category').change(function(){
		var optionSelected = $("option:selected",this).val();
		if(optionSelected === "8"){
			$('#toggleText').show();
			$('select[name=Basic_state]').prop('disabled',true);
		}
		else
		{
			$('#toggleText').hide();
			$('select[name=Basic_state]').prop('disabled',false);
		}
	});

});