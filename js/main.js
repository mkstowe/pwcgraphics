$(() => {
	var current = location.pathname;
	$("a").each(() => {
		var $this = $(this);
		if ($this.attr("href") === current) {
			$this.addClass("active");
		} else if (
			$this.attr("href") === "/index.html" &&
			current === "/"
		) {
			$this.addClass("active");
		}
	});
});

$(window).on("load", () => {
	var headerHeight = $(".header").outerHeight(true);
	var topnav = $("#topnav");
	var main = $("#main-section");
	var phantom = $("#sticky-phantom");

	$(window).resize(() => {
		headerHeight = $(".header").outerHeight(true);
	});

	$(window).bind("scroll", () => {
		if ($(window).scrollTop() > headerHeight) {
			topnav.addClass("fixed-top");
			main.css("padding-top", topnav.outerHeight(true));
			phantom.show();
		} else {
			topnav.removeClass("fixed-top");
			main.css("padding-top", 0);
			phantom.hide();
		}
	});
});

$("#preview").on("keyup", (event) => {
	$(".font-preview").html($(event.currentTarget).val());
});

// $("#bold-checkbox").change(function () {
// 	if (this.checked) {
// 		$(".font-preview").addClass("font-weight-bold");
// 	} else {
// 		$(".font-preview").removeClass("font-weight-bold");
// 	}
// });

// $("#italic-checkbox").change(function () {
// 	if (this.checked) {
// 		$(".font-preview").addClass("font-italic");
// 	} else {
// 		$(".font-preview").removeClass("font-italic");
// 	}
// });

// $("#uppercase-checkbox").change(function () {
// 	if (this.checked) {
// 		$(".font-preview").addClass("text-uppercase");
// 	} else {
// 		$(".font-preview").removeClass("text-uppercase");
// 	}
// });

$(window).on("load", () => {
	$("#preview").val("");
	// $(".checkbox").prop("checked", false);
});
