// 数値を三桁区切り表記
$("td:not(.String, .DateTime)").text(function(index, text) {
	return text.replace(/(\d)(?=(\d{3})+(?!\d))/g, '$1,');
});
 