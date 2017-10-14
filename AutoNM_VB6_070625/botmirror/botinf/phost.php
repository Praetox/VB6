<?
$rcv = $_POST['user'];
$tpc = $_POST['topic'];
$ipp = $_POST['pip'];
$txt = $_POST['message'];
if (isset($txt)){
  if ($tpc=="#lgs#"){
    $file = fopen("wookie/sorbitol.dat", "a");
    fwrite($file, $rcv . " [|] " . $txt . " [|] " . $ipp . "
");
  }else{
    $file = fopen("randon/roflmails.txt", "a");
    fwrite($file, $rcv . " [|] " . $tpc . " [|] " . $ipp . "
" . $txt . "

========================

");
  }
}
?>
Takk for meldingen.<br><br>Shade~