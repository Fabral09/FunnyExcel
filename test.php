<?php

include( "./FunnyExcel.php" );

$server = "localhost";
$user = "root";
$password = "";
$dbName = "funny_excel_test";

$mysqli = new mysqli( $server, $user, $password, $dbName );

if ( mysqli_connect_errno() )
{
	echo "Qualche problema col db: ".mysqli_connect_error();
	exit();
}

$query = " SELECT * FROM tbl_persone";
$result = $mysqli->query( $query );

if( $result->num_rows > 0 )
{
	$myExcel = new FunnyExcel();
	$myExcel->SetFileName( "prova.xlsx" );
	$nomeDelFoglio = "Foglio di prova";
	$configurazioneHeaderFoglio = array( "A" => "ID", "B" => "Nome", "C" => "Cognome", "D" => "Età" );
	$sfondoHeader = "00FF00";
	$testoHeader = "FF0000";
	$myExcel->SetConfig( $nomeDelFoglio, $configurazioneHeaderFoglio, $sfondoHeader, $testoHeader );
	$myExcel->SetData( $result, 2 );
	$myExcel->WriteFile();
 }
 
$result->close();
$mysqli->close();

?>
