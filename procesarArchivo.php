<?php
$serverName = "10.11.0.34"; //serverName\instanceName
$connectionInfo = array( "Database"=>"PPTEST", "UID"=>"carlosj", "PWD"=>"martha01");
$conn = sqlsrv_connect( $serverName, $connectionInfo);

if( $conn ) {
    //echo "Connection established.<br />";
}else{
    echo "Connection could not be established.<br />";
    die( print_r( sqlsrv_errors(), true));
}
?>