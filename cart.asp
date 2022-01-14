<!DOCTYPE html>
<html lang="en">
	<head id="head-component">
		<meta charset="UTF-8" />
	</head>

	<body>
		<div id="top-section-component"></div>

		<div class="container-fluid mt-3 px-3 px-sm-5">
			<div
				class="row d-flex justify-content-start flex-column flex-md-row"
			>
				<div
					id="sidenav-component"
					class="col-12 col-md-4 col-xl-2"
				></div>

				<div class="col-12 col-md-8 col-xl-10">
					<main class="content pl-md-3 my-3">
						<!--#include file="vsadmin/db_conn_open.asp"-->
						<!--#include file="vsadmin/inc/languagefile.asp"-->
						<!--#include file="vsadmin/includes.asp"-->
						<!--#include file="vsadmin/inc/incfunctions.asp"-->
						<!--#include file="vsadmin/inc/inccart.asp"-->
					</main>
				</div>
			</div>
		</div>

		<div id="bottom-section-component"></div>

		<script
			src="https://code.jquery.com/jquery-3.6.0.js"
			integrity="sha256-H+K7U5CnXl1h5ywQfKtSj8PCmoN9aaq30gDh27Xc0jk="
			crossorigin="anonymous"
		></script>
		<script
			src="https://cdnjs.cloudflare.com/ajax/libs/popper.js/1.12.9/umd/popper.min.js"
			integrity="sha384-ApNbgh9B+Y1QKtv3Rn7W3mgPxhU9K/ScQsAP7hUibX39j7fakFPskvXusvfa0b4Q"
			crossorigin="anonymous"
		></script>
		<script
			src="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0/js/bootstrap.min.js"
			integrity="sha384-JZR6Spejh4U02d8jOt6vLEHfe/JQGiRRSQQxSfFWpi1MquVdAyjUar5+76PVCmYl"
			crossorigin="anonymous"
		></script>
		<script src="/js/jquery.sticky.js"></script>
		<script src="/js/main.js"></script>
	</body>
</html>
