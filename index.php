<!DOCTYPE html>
<html>
<head>
	<meta charset="utf-8">
	<meta name="viewport" content="width=device-width, initial-scale=1">
	<title>Preview files with PHPWord</title>
	<link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.4.1/css/bootstrap.min.css">
</head>
<body>

	<style type="text/css">

		body{
			display: flex;
			flex-flow: column;
			justify-content: center;
			background: aliceblue;
		}

		.form{
			width: 450px;
			margin: auto;
			margin-top: 70px;
			padding: 20px;
			background: #f9f6ff;
			border:  1px solid #999;
			border-radius: 10px;
		}

		iframe{
			width: 650px;
			height: 540px;
			border: 1px solid #eee;
			outline: none;
			margin: 50px auto;
			background: #fefefe;
		}

	</style>

	<form class="form" method="post" enctype="multipart/form-data">
		<h3>Preview files with PHPWord</h3>
		<!-- https://github.com/ianconnon/ -->
		<input class="form-control" type="file" name="fileDocx"/><br>
		<input type="submit" name="send" class="btn btn-success" value="Upload Docx" type="submit"></button>
	</form>

	<?php 

	if(isset($_POST['send'])){

		require_once 'bootstrap.php';

		if(!empty($_FILES['fileDocx'])){

		    $existsFile = file_exists('previewHTML/'.$_FILES['fileDocx']['name'].'.html');


		    $mimeType = $_FILES['fileDocx']['type'];
		    $name = $_FILES['fileDocx']['name'];
		    $tmpFile = $_FILES['fileDocx']['tmp_name'];

		    if($mimeType == 'application/vnd.google-apps.document' || $mimeType == 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' || $mimeType == 'application/msword'){

		        if($existsFile == true){

		            echo '<iframe  src="previewHTML/'.$name.'.html"></iframe>';

		        }else{

		            $phpWord = \PhpOffice\PhpWord\IOFactory::load($tmpFile);
		            $htmlWriter = new \PhpOffice\PhpWord\Writer\HTML($phpWord);
		            $htmlWriter->save('previewHTML/'.$name.'.html');

		            echo '<iframe src="previewHTML/'.$name.'.html"></iframe>';

		        }

		    }else{
		        echo '<!DOCTYPE html><html><head><title>Not preview</title></head><body><p>This file hasn´t previews.</p><br>'.$mimeType.'</body></html>';
		    }

		}else{
		    echo '<!DOCTYPE html><html><head><title>Empty File</title></head><body><p>This file is empty.</p></body></html>';
		}

	}

	?>

</body>
</html>