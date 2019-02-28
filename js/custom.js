


setTimeout(function(){ 
      $('.loading').removeClass('active'); 
}, 1000);

$('.tab-header').click(function () {
    $('#tab-header-home').removeClass('accessibility-active');
    $('#tab-header-insert').removeClass('accessibility-active');
    $('#tab-header-pagelayout').removeClass('accessibility-active');
    $('#tab-header-formulas').removeClass('accessibility-active');
    $('#tab-header-data').removeClass('accessibility-active');
    $('#tab-header-review').removeClass('accessibility-active');
    $('#tab-header-view').removeClass('accessibility-active');

    $('.tab-header').removeClass('tab-header-selected');
    $(this).addClass('tab-header-selected');
});

$('#tab-header-home').click(function () {
    $('.ribbon-tab-container').hide();
    $('#ribbon-tab-container-home').show();
});

$('#tab-header-insert').click(function () {
    $('.ribbon-tab-container').hide();
    $('#ribbon-tab-container-insert').show();
});

$('#tab-header-pagelayout').click(function () {
    $('.ribbon-tab-container').hide();
    $('#ribbon-tab-container-pagelayout').show();
});

$('#tab-header-formulas').click(function () {
    $('.ribbon-tab-container').hide();
    $('#ribbon-tab-container-formulas').show();
});

$('#tab-header-data').click(function () {
    $('.ribbon-tab-container').hide();
    $('#ribbon-tab-container-data').show();
});

$('#tab-header-review').click(function () {
    $('.ribbon-tab-container').hide();
    $('#ribbon-tab-container-review').show();
});

$('#tab-header-view').click(function () {
    $('.ribbon-tab-container').hide();
    $('#ribbon-tab-container-view').show();
});


function startTime() {
    var today = new Date();
    var h = today.getHours();
    var m = today.getMinutes();
    var s = today.getSeconds();
    var mm = today.getMonth();
    var 
    m = checkTime(m);
    s = checkTime(s);
    ampm = today.getHours() >= 12 ? 'PM' : 'AM'
    months = ['1','2','3','4','5','6','7','8','9','10','11','12']
    document.getElementById('time').innerHTML =
    h + ":" + m +" "+ampm;
        

    document.getElementById('date').innerHTML =
    months[mm]+ "/" +  today.getDate() + "/" +  today.getFullYear();
    var t = setTimeout(startTime, 500);
}
function checkTime(i) {
    if (i < 10) {i = "0" + i};  // add zero in front of numbers < 10
    return i;
}


document.onkeyup = KeyCheck;       
function KeyCheck(e)
{
   var KeyID = (window.event) ? event.keyCode : e.keyCode;
  //alert(KeyID);
   if(KeyID==18){
        $('#tab-header-home').addClass('accessibility-active');
        $('#tab-header-insert').addClass('accessibility-active');
        $('#tab-header-pagelayout').addClass('accessibility-active');
        $('#tab-header-formulas').addClass('accessibility-active');
        $('#tab-header-data').addClass('accessibility-active');
        $('#tab-header-review').addClass('accessibility-active');
        $('#tab-header-view').addClass('accessibility-active');
   } else if (KeyID==27) {
        $('#tab-header-home').removeClass('accessibility-active');
        $('#tab-header-insert').removeClass('accessibility-active');
        $('#tab-header-pagelayout').removeClass('accessibility-active');
        $('#tab-header-formulas').removeClass('accessibility-active');
        $('#tab-header-data').removeClass('accessibility-active');
        $('#tab-header-review').removeClass('accessibility-active');
        $('#tab-header-view').removeClass('accessibility-active');
   }
}




// Cell Apped

$(function()  {

  $('.intExcelGrid .excel-interactor').click(function(e) {

    //getting height and width of the message box
    var height = $('#popuup_div').height();
    var width = $('#popuup_div').width();
    
    //calculating offset for displaying popup message
    leftVal=e.pageX-(width/2)+"px";
    topVal=e.pageY-(height/2)+"px";

    //show the popup message and hide with fading effect
    $('#popuup_div').empty();
    $('#popuup_div').css({left:leftVal,top:topVal}).show();

  });

 }); 


$(window).load(function(e) {
    // executes when complete page is fully loaded, including all frames, objects and images
    var top_value =  $('.excel_cell_header_bar').offset();
    var left_value =  $('.excel_cell_row_gap').offset();
  
    //alert("Top: Value" + top_value.top + " PX "  +  " Left: Value"+ left_value.left +" PX ");

});



//  TAB for sheet
$('.sheet-1-tab').click(function () {
    $('#popuup_div').hide();
    $('.sheet-1-page').show();
    $('.sheet-2-page').hide();
});

$('.sheet-2-tab').click(function () {
    $('#popuup_div').hide();
    $('.sheet-2-page').show();
    $('.sheet-1-page').hide();
}); 




$('.pmt_dialog_box').click(function () {
    $('.pmt_message_box').show();
});

$('body').on('click', '.pmt_ok_btn', function (e) {
    $('.pmt_message_box').hide();
});



$( function() {
    $( ".pmt_message_box" ).draggable();
} );

$('body').on('click', '.pmt_close', function (e) {
    $('.pmt_message_box').hide();
});
