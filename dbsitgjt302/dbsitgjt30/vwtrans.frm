TYPE=VIEW
query=(select `a`.`nomer` AS `nomer`,`a`.`barang` AS `barang`,`a`.`tglbongkar` AS `tglmulai`,`a`.`pemilik` AS `pemilik`,`a`.`tujuan` AS `tujuan`,`a`.`barge` AS `barge`,`a`.`tugboat` AS `tugboat`,`a`.`nodermaga` AS `nodermaga`,`c`.`nmpt` AS `nmtruk`,`b`.`nolambung` AS `nolambung`,`b`.`nopol` AS `nopol`,`b`.`wMasuk` AS `wmasuk`,`b`.`wKeluar` AS `wkeluar`,`b`.`bruto` AS `bruto`,`b`.`tara` AS `tara`,(`b`.`bruto` - `b`.`tara`) AS `netto`,`b`.`usergrp` AS `usergrp` from ((`dbsitgjt30`.`tbjadwal` `a` join `dbsitgjt30`.`tbtrans` `b`) join `dbsitgjt30`.`tbtruk` `c`) where ((`b`.`nomer` = `a`.`nomer`) and (`b`.`nolambung` = `c`.`nolambung`)))
md5=6c7ea6e9e4220d12c37028ac9bf9dfb4
updatable=1
algorithm=0
definer_user=root
definer_host=localhost
suid=1
with_check_option=0
revision=1
timestamp=2012-11-14 00:01:31
create-version=1
source=(select `a`.`nomer` AS `nomer`,`a`.`barang` AS `barang`,`a`.`tglbongkar` AS `tglmulai`,`a`.`pemilik` AS `pemilik`,`a`.`tujuan` AS `tujuan`,`a`.`barge` AS `barge`,`a`.`tugboat` AS `tugboat`,`a`.`nodermaga` AS `nodermaga`,`c`.`nmpt` AS `nmtruk`,\n`b`.`nolambung` AS `nolambung`,`b`.`nopol` AS `nopol`,`b`.`wMasuk` AS `wmasuk`,`b`.`wKeluar` AS `wkeluar`,`b`.`bruto` AS `bruto`,`b`.`tara` AS `tara`,(`b`.`bruto` - `b`.`tara`) AS `netto`,`b`.`usergrp` AS `usergrp` from ((`dbsitgjt30`.`tbjadwal` `a` join `dbsitgjt30`.`tbtrans` `b`) join `dbsitgjt30`.`tbtruk` `c`) where ((`b`.`nomer` = `a`.`nomer`) and (`b`.`nolambung` = `c`.`nolambung`)))
