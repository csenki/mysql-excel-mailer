CREATE TABLE `mailq` (
  `rowid` bigint(20) NOT NULL AUTO_INCREMENT,
  `msg_from` varchar(5000) COLLATE utf8_hungarian_ci NOT NULL DEFAULT 'msg@from.com',
  `msg_to` varchar(5000) COLLATE utf8_hungarian_ci NOT NULL DEFAULT 'msg@to.com',
  `msg_subject` varchar(5000) COLLATE utf8_hungarian_ci NOT NULL DEFAULT 'MSG SUBJ',
  `msg_body` mediumtext COLLATE utf8_hungarian_ci NOT NULL DEFAULT '',
  `msg_ready` tinyint(4) NOT NULL DEFAULT '0' COMMENT '2 under process 1 ready',,
  `insert_date` timestamp NULL DEFAULT CURRENT_TIMESTAMP,
  `mode` tinyint(4) NOT NULL DEFAULT '1' COMMENT '[1] Simple text mail ,[2] Do sql_command and send email with attachment ; [3] do sql_command and seve file',
  `sql_command` mediumtext COLLATE utf8_hungarian_ci NOT NULL COMMENT 'Select commands',
  `on_error` varchar(200) COLLATE utf8_hungarian_ci NOT NULL DEFAULT '',
  `file_name` varchar(200) COLLATE utf8_hungarian_ci NOT NULL DEFAULT '' COMMENT 'File name',
  `xls_opts` varchar(5000) COLLATE utf8_hungarian_ci DEFAULT '' COMMENT 'XLS parameters in JSON',
  PRIMARY KEY (`rowid`)
) ENGINE=InnoDB  DEFAULT CHARSET=utf8 COLLATE=utf8;