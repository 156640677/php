/*
SQLyog Job Agent v11.27 (32 bit) Copyright(c) Webyog Inc. All Rights Reserved.


MySQL - 5.5.47 : Database - taskms
*********************************************************************
*/

/*!40101 SET NAMES utf8 */;

/*!40101 SET SQL_MODE=''*/;

/*!40014 SET @OLD_UNIQUE_CHECKS=@@UNIQUE_CHECKS, UNIQUE_CHECKS=0 */;
/*!40014 SET @OLD_FOREIGN_KEY_CHECKS=@@FOREIGN_KEY_CHECKS, FOREIGN_KEY_CHECKS=0 */;
/*!40101 SET @OLD_SQL_MODE=@@SQL_MODE, SQL_MODE='NO_AUTO_VALUE_ON_ZERO' */;
/*!40111 SET @OLD_SQL_NOTES=@@SQL_NOTES, SQL_NOTES=0 */;
CREATE DATABASE /*!32312 IF NOT EXISTS*/`taskms` /*!40100 DEFAULT CHARACTER SET utf8 */;

USE `taskms`;

/*Table structure for table `tms_excel_config` */

DROP TABLE IF EXISTS `tms_excel_config`;

CREATE TABLE `tms_excel_config` (
  `excel_config_id` int(11) NOT NULL AUTO_INCREMENT COMMENT '主键ID',
  `excel_no` varchar(100) DEFAULT NULL COMMENT 'excel的编号',
  `excel_name` varchar(100) DEFAULT NULL COMMENT 'excel的名称',
  `begin_row` varchar(10) DEFAULT NULL COMMENT '循环部分开始行C3',
  `template_name` varchar(100) DEFAULT NULL COMMENT '模板的名称',
  `vector` varchar(10) DEFAULT 'Y' COMMENT '扩展方向Y轴方向，X轴方向,如果两个方向扩展则为XY',
  PRIMARY KEY (`excel_config_id`)
) ENGINE=MyISAM AUTO_INCREMENT=5 DEFAULT CHARSET=utf8;

/*Data for the table `tms_excel_config` */

insert  into `tms_excel_config` values (1,'1','test','2','test.xlsx','Y'),(2,'2','dv-ftp','3','dv_ftp.xlsx','Y'),(3,'3','compare_specification','B2','compare_specification.xlsx','X'),(4,'4','dashboard','A1','dashboard.xlsx','X');

/*Table structure for table `tms_excel_config_sub` */

DROP TABLE IF EXISTS `tms_excel_config_sub`;

CREATE TABLE `tms_excel_config_sub` (
  `excel_config_sub_id` int(11) NOT NULL AUTO_INCREMENT COMMENT '主键ID',
  `excel_config_id` int(11) DEFAULT NULL COMMENT 'excel的编号',
  `obj_name` varchar(100) DEFAULT NULL COMMENT '表名或者对象名',
  `col` varchar(100) DEFAULT NULL COMMENT '字段名',
  `col_name` varchar(100) DEFAULT NULL COMMENT '字段描述',
  `position` varchar(10) DEFAULT NULL COMMENT '字段位置A1、B12',
  `expression` varchar(1000) DEFAULT NULL COMMENT '表达式或者公式,行变量用#row表示，如：SUMIF($K$2:$AH$2,"Att",K#row:AH#row)，第3行时结果为：SUMIF($K$2:$AH$2,"Att",K3:AH3)，第4行时SUMIF($K$2:$AH$2,"Att",K4:AH4),只支持同一行公式，如果有col值，用#var表示',
  `is_form` int(1) DEFAULT '0' COMMENT '是否为行配置，0不是，1是',
  PRIMARY KEY (`excel_config_sub_id`)
) ENGINE=MyISAM AUTO_INCREMENT=20 DEFAULT CHARSET=utf8;

/*Data for the table `tms_excel_config_sub` */

insert  into `tms_excel_config_sub` values (1,1,'1','col1','col1','A',NULL,0),(2,1,'2','col2','col2','B',NULL,0),(3,1,'3','col3','col3','C','CONCATENATE(A#row+B#row,#var)',0),(4,2,'1','col1','col1','A',NULL,0),(5,2,'2','col2','col2','B',NULL,0),(6,2,'3','col3','col3','C',NULL,0),(7,2,'4','col4','col4','G','SUMIF($K$2:$AH$2,\"Att\",K#row:AH#row)',0),(8,2,'5','col5','col5','I','G3-H3',0),(9,2,'6','col6','col6','J','H3/G3',0),(10,2,'7','col7','col7','K',NULL,0),(11,2,'8','col8','col8','L',NULL,0),(12,2,'9','col9','col9','M',NULL,0),(13,2,'10','col10','col10','N',NULL,0),(15,2,'12','col12','col12','H','SUMIF($K$2:$AH$2,\"Pass\",K#row:AH#row)',0),(17,2,'13','form1','form1','K1',NULL,1),(18,2,'14','form2','form2','M1',NULL,1);

/*!40101 SET SQL_MODE=@OLD_SQL_MODE */;
/*!40014 SET FOREIGN_KEY_CHECKS=@OLD_FOREIGN_KEY_CHECKS */;
/*!40014 SET UNIQUE_CHECKS=@OLD_UNIQUE_CHECKS */;
/*!40111 SET SQL_NOTES=@OLD_SQL_NOTES */;
