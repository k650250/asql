// ���l���O����؂�\�L
$("td:not(.String, .DateTime)").text(function(index, text) {
	return text.replace(/(\d)(?=(\d{3})+(?!\d))/g, '$1,');
});
 