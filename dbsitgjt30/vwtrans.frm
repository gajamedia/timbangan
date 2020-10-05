TYPE=VIEW
query=(select `a`.`nomer` AS `nomer`,`a`.`barang` AS `barang`,`a`.`tglbongkar` AS `tglmulai`,`a`.`pemilik` AS `pemilik`,`a`.`tujuan` AS `tujuan`,`a`.`barge` AS `barge`,`a`.`tugboat` AS `tugboat`,`a`.`nodermaga` AS `nodermaga`,`c`.`nmpt` AS `nmtruk`,`b`.`nolambung` AS `nolambung`,`b`.`nopol` AS `nopol`,`b`.`wMasuk` AS `wmasuk`,`b`.`wKeluar` AS `wkeluar`,`b`.`bruto` AS `bruto`,`b`.`tara` AS `tara`,(`b`.`bruto` - `b`.`tara`) AS `netto`,`b`.`usergrp` AS `usergrp`,`a`.`bl` AS `bl` from ((`dbsitgjt30`.`tbjadwal` `a` join `dbsitgjt30`.`tbtrans` `b`) join `dbsitgjt30`.`tbtruk` `c`) where ((`b`.`nomer` = `a`.`nomer`) and (`b`.`nolambung` = `c`.`nolambung`)))
md5=ad292767d5517a81bffbb027956e0892
updatable=1
algorithm=0
definer_user=root
definer_host=localhost
suid=1
with_check_option=0
revision=1
timestamp=2013-09-05 19:43:43
create-version=1
source=(\nselect\n  `a`.`nomer`      AS `nomer`,\n  `a`.`barang`     AS `barang`,\n  `a`.`tglbongkar` AS `tglmulai`,\n  `a`.`pemilik`    AS `pemilik`,\n  `a`.`tujuan`     AS `tujuan`,\n  `a`.`barge`      AS `barge`,\n  `a`.`tugboat`    AS `tugboat`,\n  `a`.`nodermaga`  AS `nodermaga`,\n  `c`.`nmpt`       AS `nmtruk`,\n  `b`.`nolambung`  AS `nolambung`,\n  `b`.`nopol`      AS `nopol`,\n  `b`.`wMasuk`     AS `wmasuk`,\n  `b`.`wKeluar`    AS `wkeluar`,\n  `b`.`bruto`      AS `bruto`,\n  `b`.`tara`       AS `tara`,\n  (`b`.`bruto` - `b`.`tara`) AS `netto`,\n  `b`.`usergrp`    AS `usergrp`,\n  `a`.`bl`         AS `bl`\nfrom ((`tbjadwal` `a`\n    join `tbtrans` `b`)\n   join `tbtruk` `c`)\nwhere ((`b`.`nomer` = `a`.`nomer`)\n       and (`b`.`nolambung` = `c`.`nolambung`)))
