// JavaScript Document
$(document).ready(function(){
//superfish下拉選單
$("#mainmenu").superfish({
	delay:         300,
	});	
});
//修正垂直easytab bug
var tabVheight = $(".easytabV ul.menubar").height();
var tabVwidth = $(".easytabV .menubar li").width();
var tabpanelwidth = $(".easytabV .panel-container").width();
$(".easytabV .panel-container").css("width", tabpanelwidth - tabVwidth + "px");
$(".easytabV .panel-container").css("margin-left", tabVwidth + "px");
$(".easytabV .panel-container>div").css("min-height", tabVheight + "px");
//colorbox
	$(".colorboxS").colorbox({inline:true, width:"640", maxHeight:"480", opacity:0.5});
	$(".colorboxM").colorbox({inline:true, width:"800", height:"600", opacity:0.5});
	$(".colorboxL").colorbox({inline:true, width:"900" ,opacity:0.5});
	$(".colorboxGen").colorbox({inline:true, width:"80%",maxWidth:"980", MaxHeight:"80%" ,opacity:0.5});
	$("#colorboxiframe").colorbox({iframe:true, width:"80%",maxWidth:"980", MaxHeight:"80%",opacity:0.5});
	//修改外框
	$("#cboxTopLeft").hide();
	$("#cboxTopRight").hide();
	$("#cboxBottomLeft").hide();
	$("#cboxBottomRight").hide();
	$("#cboxMiddleLeft").hide();
	$("#cboxMiddleRight").hide();
	$("#cboxTopCenter").hide();
	$("#cboxBottomCenter").hide();
	$("#cboxContent").addClass("colorboxnewborder");
//絕對置底
$(document).ready(sizeContent);
$(window).resize(sizeContent);

function sizeContent() {
	var overHeight = $(".WrapperBody").height();// 內容大於視窗高度
    var newHeight = $("html").height() - $(".WrapperFooter").height();//內容小於視窗高度
if(overHeight > newHeight){
	$(".WrapperBody").css("min-height", overHeight + "px");
	}
	else{
	$(".WrapperBody").css("min-height", newHeight + "px");
	}
}