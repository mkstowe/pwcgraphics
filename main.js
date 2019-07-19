// Change Facebook icon color on hover
document.getElementById("facebook").onmouseover = function () {
    document.getElementById("fb-circle").style.color = "#414141";
};

document.getElementById("facebook").onmouseleave = function () {
    document.getElementById("fb-circle").style.color = "#2D2D2D";
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
