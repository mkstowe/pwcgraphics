// Change Facebook icon color on hover
document.getElementById("facebook").onmouseover = function () {
    document.getElementById("fb-circle").style.color = "#FCC605";
};

document.getElementById("facebook").onmouseleave = function () {
    document.getElementById("fb-circle").style.color = "#F3F3F3";
};

// Change Twitter icon color on hover
document.getElementById("twitter").onmouseover = function () {
    document.getElementById("twitter-circle").style.color = "#FCC605";
};

document.getElementById("twitter").onmouseleave = function () {
    document.getElementById("twitter-circle").style.color = "#F3F3F3";
};

// Change Instagram icon color on hover
document.getElementById("insta").onmouseover = function () {
    document.getElementById("ig-circle").style.color = "#FCC605";
};

document.getElementById("insta").onmouseleave = function () {
    document.getElementById("ig-circle").style.color = "#F3F3F3";
};

// Change cart icon on hover
document.getElementById("cart").onmouseover = function () {
    this.style.color = "#FCC605";
};

document.getElementById("cart").onmouseleave = function () {
    this.style.color = "#F3F3F3";
};

$(function () {
    $("#navbar").sticky({topSpacing: 0});
});

function openTopMenu() {
    if (window.innerWidth < 680) {
        document.getElementById("dropdown-content").style.display = "block";
    }
}

window.onclick = function (event) {
    if (window.innerWidth < 680) {
        if (event.target !== document.getElementById("dropbtn")) {
            document.getElementById("dropdown-content").style.display = "none";
        }
    }
};
