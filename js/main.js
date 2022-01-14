$("#topnav").sticky({ topSpacing: 0 });

$(function () {
	$("#head-component").load("/components/head-component.html");
	$("#top-section-component").load(
		"/components/top-section-component.html",
		function () {
			var current = location.pathname;
			$("a").each(function () {
				var $this = $(this);
				if ($this.attr("href").indexOf(current) !== -1) {
					$this.addClass("active");
				}
			});
		}
	);
	$("#sidenav-component").load(
		"/components/sidenav-component.html",
		function () {
			var current = location.pathname;
			$("a").each(function () {
				var $this = $(this);
				if ($this.attr("href").indexOf(current) !== -1) {
					$this.addClass("active");
				}
			});
		}
	);
	$("#bottom-section-component").load(
		"/components/bottom-section-component.html"
	);
});
