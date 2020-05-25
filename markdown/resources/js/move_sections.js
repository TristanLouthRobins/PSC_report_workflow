$(document).ready(function(){
 $("#introduction").children()[0].textContent = "";

 $("#introduction").children().addClass("toc-ignore");

 $("#introduction").children().appendTo($("#new_intro"));

 $("#introduction").remove();


});
