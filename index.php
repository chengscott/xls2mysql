<?php
require_once 'Excel/reader.php';
set_time_limit(10000);
$data = new Spreadsheet_Excel_Reader();
$data->setOutputEncoding('UTF-8');
$data->setUTFEncoder('mb');//use 'mb_convert_encoding' to replace 'iconv'
$data->read('student.xls');

$mysqli = new mysqli('localhost', 'UserBlah', 'blablabla', 'mydb');
$mysqli->set_charset('utf8');
if (mysqli_connect_errno()) {
  die("Connect failed ({$mysqli->connect_errno}): {$mysqli->connect_error}");
}

$col = array('', '', '', '');

$sql = "INSERT INTO student SET sid=?, class=?, seatnumber=?, name =?;";
if ($stmt = $mysqli->prepare($sql)) {
  $stmt->bind_param("iiis", $col[1], $col[2], $col[3], $col[4]);
} else {
  echo 'Server Error';
}


$numRows = $data->sheets[0]['numRows'];
$numCols = $data->sheets[0]['numCols'];

for ($i = 1; $i <= $numRows; ++$i) {
  for ($j = 1; $j <= $numCols; ++$j) {
    $col[$j] = $data->sheets[0]['cells'][$i][$j];
  }
  $stmt->execute();
  echo "Processed {$i} rows.<br>";
}
echo "---<br>Finished! ({$numRows} rows.)";
$stmt->close();
$mysqli->close();
?>