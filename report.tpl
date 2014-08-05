<html>
	<head>
		<title>Validate Hyperlinks in Word Documents</title>
	</head>
	<body>
		<div class="container">

			<h1>Directories checked</h1>

			<ul>
				<li>sample directory</li>
			</ul>

			<h1>Report</h1>

			<div class="result">
				All Files are valid
			</div>

			<div class="legend">
				<p>show only invalid links</p>
				<p>show all links</p>
			</div>

			<ul>
				{{range .}}
				<li class="result">
					<h2>{{.Path}}</h2>

					<ul>
						{{range .Hyperlinks}}
						<li class="result">{{.}}</li>
						{{end}}
					</ul>
				</li>
				{{end}}
			</ul>

		</div>
	</body>
</html>
