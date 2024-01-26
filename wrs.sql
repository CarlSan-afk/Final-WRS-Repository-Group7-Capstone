-- MySQL dump 10.13  Distrib 5.6.17, for Win64 (x86_64)
--
-- Host: localhost    Database: wrs
-- ------------------------------------------------------
-- Server version	5.5.41

/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8 */;
/*!40103 SET @OLD_TIME_ZONE=@@TIME_ZONE */;
/*!40103 SET TIME_ZONE='+00:00' */;
/*!40014 SET @OLD_UNIQUE_CHECKS=@@UNIQUE_CHECKS, UNIQUE_CHECKS=0 */;
/*!40014 SET @OLD_FOREIGN_KEY_CHECKS=@@FOREIGN_KEY_CHECKS, FOREIGN_KEY_CHECKS=0 */;
/*!40101 SET @OLD_SQL_MODE=@@SQL_MODE, SQL_MODE='NO_AUTO_VALUE_ON_ZERO' */;
/*!40111 SET @OLD_SQL_NOTES=@@SQL_NOTES, SQL_NOTES=0 */;

--
-- Table structure for table `tbl_account`
--

DROP TABLE IF EXISTS `tbl_account`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `tbl_account` (
  `username` varchar(45) DEFAULT NULL,
  `password` varchar(45) DEFAULT NULL,
  `id` int(11) NOT NULL AUTO_INCREMENT,
  `role` varchar(45) DEFAULT NULL,
  PRIMARY KEY (`id`)
) ENGINE=InnoDB AUTO_INCREMENT=11 DEFAULT CHARSET=latin1;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `tbl_account`
--

LOCK TABLES `tbl_account` WRITE;
/*!40000 ALTER TABLE `tbl_account` DISABLE KEYS */;
INSERT INTO `tbl_account` VALUES ('admin','qWq',5,'user'),('cash1','321',6,'user'),('X','1',10,'admin');
/*!40000 ALTER TABLE `tbl_account` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `tbl_credit`
--

DROP TABLE IF EXISTS `tbl_credit`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `tbl_credit` (
  `Balance` decimal(10,0) DEFAULT NULL,
  `amount` decimal(10,0) DEFAULT NULL,
  `Customer_name` varchar(100) DEFAULT NULL,
  `classification` varchar(45) DEFAULT NULL,
  `Id_number` int(11) DEFAULT NULL,
  `date_of_sale` datetime DEFAULT NULL,
  `delivered_by` varchar(45) DEFAULT NULL,
  `Sold_item` varchar(200) DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `tbl_credit`
--

LOCK TABLES `tbl_credit` WRITE;
/*!40000 ALTER TABLE `tbl_credit` DISABLE KEYS */;
/*!40000 ALTER TABLE `tbl_credit` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `tbl_customer_info`
--

DROP TABLE IF EXISTS `tbl_customer_info`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `tbl_customer_info` (
  `ID_number` int(11) NOT NULL AUTO_INCREMENT,
  `Customer_Name` varchar(45) DEFAULT NULL,
  `Classification` varchar(45) DEFAULT NULL,
  `Address` varchar(100) DEFAULT NULL,
  `facebook` varchar(45) DEFAULT NULL,
  `date_of_last_buy` datetime DEFAULT NULL,
  `image` longtext,
  `contact` mediumtext,
  PRIMARY KEY (`ID_number`)
) ENGINE=InnoDB AUTO_INCREMENT=1006 DEFAULT CHARSET=latin1;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `tbl_customer_info`
--

LOCK TABLES `tbl_customer_info` WRITE;
/*!40000 ALTER TABLE `tbl_customer_info` DISABLE KEYS */;
INSERT INTO `tbl_customer_info` VALUES (1000,'Guest','Household',NULL,NULL,'2023-11-11 15:42:08',NULL,NULL),(1001,'jose','Household','batobato st','fb.com','2023-10-29 17:22:03','','0909'),(1002,'rizal','Household','biak-na-bato','fb.com/rizal','2023-11-11 16:06:24','C:\\Users\\Carl\\Desktop\\top25animecharacters-blogroll-1660777571580.jpg','991111'),(1004,'test','Reseller','fdsfdsfdsferwrew','test','2023-10-14 00:00:00','C:\\Users\\AYC\\Pictures\\0-02-06-eab6ca30d46b1f2cdedcc20bff73cc30c4230b6a5507fd13c43910d2d7fc06c7_82e8cc39e458df67.jpg','ewq432432'),(1005,'nene','Household','xyz','nenefb','2023-11-03 17:22:36','','09473434343');
/*!40000 ALTER TABLE `tbl_customer_info` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `tbl_customer_item`
--

DROP TABLE IF EXISTS `tbl_customer_item`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `tbl_customer_item` (
  `ID_number` int(11) DEFAULT NULL,
  `Customer_name` varchar(100) DEFAULT NULL,
  `Item_title` varchar(100) DEFAULT NULL,
  `wrs_gallon` int(11) DEFAULT NULL,
  `date_borrowed` datetime DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `tbl_customer_item`
--

LOCK TABLES `tbl_customer_item` WRITE;
/*!40000 ALTER TABLE `tbl_customer_item` DISABLE KEYS */;
INSERT INTO `tbl_customer_item` VALUES (1004,'test','XXXX',2,'2023-10-14 00:00:00'),(1001,'jose','Slim Container with Fauset',1,'2023-10-14 00:00:00'),(1002,'rizal','Slim Container with Fauset',2,'2023-11-03 00:00:00');
/*!40000 ALTER TABLE `tbl_customer_item` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `tbl_customization`
--

DROP TABLE IF EXISTS `tbl_customization`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `tbl_customization` (
  `logo` longtext,
  `station_name` longtext
) ENGINE=InnoDB DEFAULT CHARSET=latin1;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `tbl_customization`
--

LOCK TABLES `tbl_customization` WRITE;
/*!40000 ALTER TABLE `tbl_customization` DISABLE KEYS */;
/*!40000 ALTER TABLE `tbl_customization` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `tbl_delivery`
--

DROP TABLE IF EXISTS `tbl_delivery`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `tbl_delivery` (
  `station_name` decimal(10,0) DEFAULT NULL,
  `amount` decimal(10,0) DEFAULT NULL,
  `Customer_name` varchar(45) DEFAULT NULL,
  `classification` varchar(45) DEFAULT NULL,
  `Id_number` int(11) DEFAULT NULL,
  `date_of_sale` datetime DEFAULT NULL,
  `delivered_by` varchar(45) DEFAULT NULL,
  `Sold_item` longtext,
  `TN` int(11) NOT NULL AUTO_INCREMENT,
  PRIMARY KEY (`TN`)
) ENGINE=InnoDB AUTO_INCREMENT=24 DEFAULT CHARSET=latin1;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `tbl_delivery`
--

LOCK TABLES `tbl_delivery` WRITE;
/*!40000 ALTER TABLE `tbl_delivery` DISABLE KEYS */;
/*!40000 ALTER TABLE `tbl_delivery` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `tbl_deliveryman`
--

DROP TABLE IF EXISTS `tbl_deliveryman`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `tbl_deliveryman` (
  `Deliveryman_name` varchar(100) DEFAULT NULL,
  `Contact_number` varchar(45) DEFAULT NULL,
  `Address` longtext
) ENGINE=InnoDB DEFAULT CHARSET=latin1;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `tbl_deliveryman`
--

LOCK TABLES `tbl_deliveryman` WRITE;
/*!40000 ALTER TABLE `tbl_deliveryman` DISABLE KEYS */;
INSERT INTO `tbl_deliveryman` VALUES ('carl','1111','111');
/*!40000 ALTER TABLE `tbl_deliveryman` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `tbl_expenses`
--

DROP TABLE IF EXISTS `tbl_expenses`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `tbl_expenses` (
  `Item_title` varchar(50) DEFAULT NULL,
  `Item_cost` decimal(10,0) DEFAULT NULL,
  `date_of_expenses` datetime DEFAULT NULL,
  `TN` int(11) DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `tbl_expenses`
--

LOCK TABLES `tbl_expenses` WRITE;
/*!40000 ALTER TABLE `tbl_expenses` DISABLE KEYS */;
INSERT INTO `tbl_expenses` VALUES ('Slim Container with Fauset',100000,'2023-10-08 00:00:00',NULL),('Slim Container with Fauset',36000,'2023-10-08 00:00:00',NULL),('Slim Container with Fauset',0,'2023-10-08 00:00:00',NULL),('Slim Container with Fauset',9000,'2023-10-14 00:00:00',NULL),('water dispenser',9000,'2023-10-14 00:00:00',NULL),('XXXX',200,'2023-10-14 00:00:00',NULL),('wilkins 350',100000,'2023-10-31 00:00:00',NULL),('summit 1l',99990,'2023-10-31 00:00:00',NULL),('testQR',0,'2023-11-03 00:00:00',NULL),('testQR',0,'2023-11-03 00:00:00',NULL),('testQR',0,'2023-11-03 00:00:00',NULL),('testQR',0,'2023-11-03 00:00:00',NULL),('testQR',0,'2023-11-03 00:00:00',NULL),('galon 10',100000,'2023-11-03 00:00:00',NULL),('testQR',0,'2023-11-03 00:00:00',NULL);
/*!40000 ALTER TABLE `tbl_expenses` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `tbl_history`
--

DROP TABLE IF EXISTS `tbl_history`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `tbl_history` (
  `customer_name` varchar(100) DEFAULT NULL,
  `id_number` varchar(45) DEFAULT NULL,
  `classification` varchar(45) DEFAULT NULL,
  `transaction_date` datetime DEFAULT NULL,
  `transaction_details` longtext
) ENGINE=InnoDB DEFAULT CHARSET=latin1;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `tbl_history`
--

LOCK TABLES `tbl_history` WRITE;
/*!40000 ALTER TABLE `tbl_history` DISABLE KEYS */;
INSERT INTO `tbl_history` VALUES ('Guest','1000','Household','2023-11-03 15:27:32','Deliver (1) Slim Container with Fauset / Deliver (1) testQR / '),('rizal','1002','Household','2023-11-03 15:33:45','Deliver (1) Slim Container with Fauset / Deliver (1) testQR / /Pay 57'),('nene','1005','Household','2023-11-03 15:33:55','Deliver (1) Slim Container with Fauset / Deliver (1) testQR / /Pay 57'),('Guest','1000','Household','2023-11-03 15:33:56','Buy (2) Slim Container with Fauset / Buy (1) testQR / /Pay 432'),('Guest','1000','Household','2023-11-03 15:37:26','Buy (1) Slim Container with Fauset / Buy (1) testQR / /Pay 232'),('Guest','1000','Household','2023-11-03 15:44:04','Buy (1) Slim Container with Fauset / Buy (1) XXXX / /Pay 200'),('Administrator','9999','N/A','2023-11-03 15:44:53','Added (5) Stocks of testQR to Itemlist'),('rizal','1002','Household','2023-11-03 15:46:37','Deliver (2) Slim Container with Fauset / /Pay 50'),('Guest','1000','Household','2023-11-03 15:53:21','Deliver (1) Slim Container with Fauset / /Pay 25'),('rizal','1002','N/A','2023-11-03 15:55:38','Borrowed  (2) Slim Container with Fauset'),('rizal','1002','Household','2023-11-03 16:13:21','Deliver (1) testQR / /Pay 32'),('Administrator','9999','N/A','2023-11-03 16:24:07','Added (1000) Stocks of galon 10 to Itemlist'),('Guest','1000','Household','2023-11-03 16:28:36','Buy (1) Slim Container with Fauset / Buy (1) testQR / /Pay 232'),('Guest','1000','Household','2023-11-03 17:06:09','Buy (1) Slim Container with Fauset / /Pay 200'),('Guest','1000','Household','2023-11-03 17:09:29','Buy (1) Slim Container with Fauset / Buy (1) testQR / /Pay 232'),('Guest','1000','Household','2023-11-03 17:10:30','Buy (1) Slim Container with Fauset / Buy (1) testQR / /Pay 232'),('Guest','1000','Household','2023-11-03 17:19:42','Buy (5) Slim Container with Fauset / Buy (2) testQR / /Pay 1064'),('nene','1005','Household','2023-11-03 17:22:32','Buy (1) testQR / /Pay 32'),('Guest','1000','Household','2023-11-03 17:22:53','Buy (5) Slim Container with Fauset / Buy (2) testQR / /Pay 1064'),('Administrator','9999','N/A','2023-11-03 17:23:11','Added (1000) Stocks of testQR to Itemlist'),('Guest','1000','Household','2023-11-03 17:25:18','Buy (10) Slim Container with Fauset / Buy (5) testQR / /Pay 2160'),('Administrator','9999','N/A','2023-11-11 15:27:14','Deleted Deliveryman'),('sss','1005','Household','2023-11-11 15:30:23',' Added New Customer '),('Administrator','9999','N/A','2023-11-11 15:31:28','Added New Deliveryman - carl'),('Administrator','9999','N/A','2023-11-11 15:33:57','Added New Deliveryman - carl'),('Guest','1000','Household','2023-11-11 15:34:32','Canceled Delivery with 0 Balance'),('Guest','1000','Household','2023-11-11 15:42:06','Deliver (1) Slim Container with Fauset / /Pay 25'),('rizal','1002','Household','2023-11-11 15:43:18','Canceled Delivery with -5 Balance'),('rizal','1002','Household','2023-11-11 15:46:30','Deliver (1) Slim Container with Fauset / /Pay 500'),('rizal','1002','Household','2023-11-11 15:59:38','Credited -25Amounting to -25'),('rizal','1002','Household','2023-11-11 16:01:14','Deliver (1) Slim Container with Fauset / /Pay 500'),('rizal','1002','Household','2023-11-11 16:07:05','Deliver (1) Slim Container with Fauset / /Pay 3'),('rizal','1002','Household','2023-11-11 16:08:13','Deliver (1) Slim Container with Fauset / /Pay 1000'),('Administrator','9999','N/A','2023-11-11 16:16:20','Added  3232 to Itemlist'),('Administrator','9999','N/A','2023-11-11 16:52:26','Added  RAmdonSSSS to Itemlist'),('Administrator','9999','N/A','2023-11-11 16:52:55','Edited   RAmdonSSSS'),('Administrator','9999','N/A','2023-11-11 16:53:15','Edited   RAmdonSSSS'),('Administrator','9999','N/A','2023-11-11 16:53:57','Edited   RAmdonSSSS'),('Administrator','9999','N/A','2023-11-11 16:54:14','Edited   RAmdonSSSS'),('Administrator','9999','N/A','2023-11-11 16:55:56','Edited   RAmdonSSSS');
/*!40000 ALTER TABLE `tbl_history` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `tbl_item_list`
--

DROP TABLE IF EXISTS `tbl_item_list`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `tbl_item_list` (
  `Item_title` varchar(50) DEFAULT NULL,
  `type` varchar(45) DEFAULT NULL,
  `pos_item` varchar(45) DEFAULT NULL,
  `Household_deliver_charge` decimal(10,0) DEFAULT NULL,
  `Household_pickup_charge` decimal(10,0) DEFAULT NULL,
  `Household_purchase_charge` decimal(10,0) DEFAULT NULL,
  `Reseller_deliver_charge` decimal(10,0) DEFAULT NULL,
  `Reseller_pickup_charge` decimal(10,0) DEFAULT NULL,
  `Reseller_purchase_charge` decimal(10,0) DEFAULT NULL,
  `Dealer_deliver_charge` decimal(10,0) DEFAULT NULL,
  `Dealer_pickup_charge` decimal(10,0) DEFAULT NULL,
  `tbl_item_listcol` decimal(10,0) DEFAULT NULL,
  `discription` longtext,
  `Stocks` int(11) DEFAULT '0',
  `Borrowed` int(11) DEFAULT '0',
  `Damaged` int(11) DEFAULT '0',
  `qrCode` varchar(255) DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `tbl_item_list`
--

LOCK TABLES `tbl_item_list` WRITE;
/*!40000 ALTER TABLE `tbl_item_list` DISABLE KEYS */;
INSERT INTO `tbl_item_list` VALUES ('Slim Container with Fauset','Container','Yes',25,20,200,25,20,200,20,20,200,'Default Value',9762,3,0,'3123123213'),('galon 10','Bottle','Yes',10,5,100,0,0,0,0,0,0,'',999,0,0,NULL),('test','Container','Yes',1,2,10,0,0,0,0,0,0,'',0,0,0,NULL),('water dispenser','Others','Yes',15,0,2000,20,0,2100,50,0,2150,'water dispenser mmmm',2,0,0,'7777777'),('dagdag','Container','Yes',2,5,50,2,5,44,2,5,30,'pandagdag',0,0,0,NULL),('XXXX','Container','Yes',0,0,0,0,0,0,0,0,0,'sadsad',17,2,0,NULL),('testQR','Container','Yes',32,32,32,40,0,50,323,0,32323,'32323232323',770,0,3,'3243243243233232'),('sddfsfdsf','Container','Yes',1,0,1,1,0,11,1,0,1,'fdsf',0,0,0,'32432432432'),('wilkins 350','Container','Yes',20,10,10,20,10,10,0,0,0,'',10000,0,0,'1000986789864323'),('summit 1l','Container','Yes',30,10,10,0,0,0,0,0,0,'fdf',9999,0,0,'10000232487328472'),('3232','Container','Yes',23,32,23,0,0,0,0,0,0,'3232',0,0,0,''),('RAmdonSSSS','Container','Yes',1,1,1,0,0,0,0,0,0,'32131',0,0,0,'7203351538147');
/*!40000 ALTER TABLE `tbl_item_list` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `tbl_refilled`
--

DROP TABLE IF EXISTS `tbl_refilled`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `tbl_refilled` (
  `Refilled_gallon` int(11) DEFAULT NULL,
  `Date_refilled` datetime DEFAULT NULL,
  `Id_number` int(11) DEFAULT NULL,
  `Customer_name` longtext
) ENGINE=InnoDB DEFAULT CHARSET=latin1;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `tbl_refilled`
--

LOCK TABLES `tbl_refilled` WRITE;
/*!40000 ALTER TABLE `tbl_refilled` DISABLE KEYS */;
INSERT INTO `tbl_refilled` VALUES (1,'2023-10-08 00:00:00',1000,'Guest'),(4,'2023-10-08 00:00:00',1000,'Guest'),(1,'2023-10-08 00:00:00',1000,'nene'),(1,'2023-10-08 00:00:00',1000,'nene'),(8,'2023-10-08 00:00:00',1000,'nene'),(1,'2023-10-08 00:00:00',1000,'Guest'),(2,'2023-10-08 00:00:00',1000,'nene'),(3,'2023-10-08 00:00:00',1002,'rizal'),(1,'2023-10-14 00:00:00',1002,'rizal'),(2,'2023-10-29 00:00:00',1000,'Guest'),(1,'2023-10-29 00:00:00',1002,'rizal'),(1,'2023-10-29 00:00:00',1004,'test'),(2,'2023-10-29 00:00:00',1001,'jose'),(1,'2023-10-29 00:00:00',1000,'Guest'),(2,'2023-10-29 00:00:00',1002,'rizal'),(4,'2023-10-29 00:00:00',1000,'nene'),(1,'2023-10-29 00:00:00',1001,'jose'),(1,'2023-10-29 00:00:00',1000,'Guest'),(2,'2023-10-31 00:00:00',1000,'Guest'),(2,'2023-10-31 00:00:00',1002,'rizal'),(22,'2023-11-03 00:00:00',1002,'rizal'),(2,'2023-11-03 00:00:00',1002,'rizal'),(2,'2023-11-03 00:00:00',1005,'nene'),(3,'2023-11-03 00:00:00',1000,'Guest'),(2,'2023-11-03 00:00:00',1000,'Guest'),(2,'2023-11-03 00:00:00',1000,'Guest'),(2,'2023-11-03 00:00:00',1002,'rizal'),(1,'2023-11-03 00:00:00',1000,'Guest'),(1,'2023-11-03 00:00:00',1002,'rizal'),(2,'2023-11-03 00:00:00',1000,'Guest'),(1,'2023-11-03 00:00:00',1005,'nene'),(1,'2023-11-03 00:00:00',1000,'Guest'),(2,'2023-11-03 00:00:00',1000,'Guest'),(2,'2023-11-03 00:00:00',1000,'Guest'),(7,'2023-11-03 00:00:00',1000,'Guest'),(7,'2023-11-03 00:00:00',1000,'Guest'),(15,'2023-11-03 00:00:00',1000,'Guest'),(19,'2023-11-03 00:00:00',1000,'Guest'),(1,'2023-11-11 00:00:00',1000,'Guest'),(1,'2023-11-11 00:00:00',1002,'rizal'),(1,'2023-11-11 00:00:00',1002,'rizal'),(1,'2023-11-11 00:00:00',1002,'rizal'),(1,'2023-11-11 00:00:00',1002,'rizal');
/*!40000 ALTER TABLE `tbl_refilled` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `tbl_sales`
--

DROP TABLE IF EXISTS `tbl_sales`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `tbl_sales` (
  `amount` decimal(10,0) DEFAULT NULL,
  `Customer_name` varchar(45) DEFAULT NULL,
  `classification` varchar(45) DEFAULT NULL,
  `Id_number` int(11) DEFAULT NULL,
  `date_of_sale` datetime DEFAULT NULL,
  `delivered_by` longtext,
  `Sold_item` longtext,
  `TN` int(11) NOT NULL AUTO_INCREMENT,
  PRIMARY KEY (`TN`)
) ENGINE=InnoDB AUTO_INCREMENT=23 DEFAULT CHARSET=latin1;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `tbl_sales`
--

LOCK TABLES `tbl_sales` WRITE;
/*!40000 ALTER TABLE `tbl_sales` DISABLE KEYS */;
INSERT INTO `tbl_sales` VALUES (-57,'Guest','Household',1000,'2023-11-03 00:00:00',NULL,'Deliver (1) Slim Container with Fauset / Deliver (1) testQR / ',1),(57,'rizal','Household',1002,'2023-11-03 00:00:00','N/A','Deliver (1) Slim Container with Fauset / Deliver (1) testQR / ',2),(57,'nene','Household',1005,'2023-11-03 00:00:00','N/A','Deliver (1) Slim Container with Fauset / Deliver (1) testQR / ',3),(432,'Guest','Household',1000,'2023-11-03 00:00:00','N/A','Buy (2) Slim Container with Fauset / Buy (1) testQR / ',4),(232,'Guest','Household',1000,'2023-11-03 00:00:00','N/A','Buy (1) Slim Container with Fauset / Buy (1) testQR / ',5),(200,'Guest','Household',1000,'2023-11-03 00:00:00','N/A','Buy (1) Slim Container with Fauset / Buy (1) XXXX / ',6),(50,'rizal','Household',1002,'2023-11-03 00:00:00','jose laurel','Deliver (2) Slim Container with Fauset / ',7),(25,'Guest','Household',1000,'2023-11-03 00:00:00','jose laurel','Deliver (1) Slim Container with Fauset / ',8),(32,'rizal','Household',1002,'2023-11-03 00:00:00','jose laurel','Deliver (1) testQR / ',9),(232,'Guest','Household',1000,'2023-11-03 00:00:00','jose laurel','Buy (1) Slim Container with Fauset / Buy (1) testQR / ',10),(200,'Guest','Household',1000,'2023-11-03 00:00:00','N/A','Buy (1) Slim Container with Fauset / ',11),(232,'Guest','Household',1000,'2023-11-03 00:00:00','N/A','Buy (1) Slim Container with Fauset / Buy (1) testQR / ',12),(232,'Guest','Household',1000,'2023-11-03 00:00:00','N/A','Buy (1) Slim Container with Fauset / Buy (1) testQR / ',13),(1064,'Guest','Household',1000,'2023-11-03 00:00:00','N/A','Buy (5) Slim Container with Fauset / Buy (2) testQR / ',14),(32,'nene','Household',1005,'2023-11-03 00:00:00','N/A','Buy (1) testQR / ',15),(1064,'Guest','Household',1000,'2023-11-03 00:00:00','N/A','Buy (5) Slim Container with Fauset / Buy (2) testQR / ',16),(2160,'Guest','Household',1000,'2023-11-03 00:00:00','N/A','Buy (10) Slim Container with Fauset / Buy (5) testQR / ',17),(25,'Guest','Household',1000,'2023-11-11 00:00:00','carl','Deliver (1) Slim Container with Fauset / ',18),(500,'rizal','Household',1002,'2023-11-11 00:00:00','carl','Deliver (1) Slim Container with Fauset / ',19),(500,'rizal','Household',1002,'2023-11-11 00:00:00','carl','Deliver (1) Slim Container with Fauset / ',20),(3,'rizal','Household',1002,'2023-11-11 00:00:00','carl','Deliver (1) Slim Container with Fauset / ',21),(1000,'rizal','Household',1002,'2023-11-11 00:00:00','carl','Deliver (1) Slim Container with Fauset / ',22);
/*!40000 ALTER TABLE `tbl_sales` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `tbl_systemlog`
--

DROP TABLE IF EXISTS `tbl_systemlog`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `tbl_systemlog` (
  `LogDate` datetime DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `tbl_systemlog`
--

LOCK TABLES `tbl_systemlog` WRITE;
/*!40000 ALTER TABLE `tbl_systemlog` DISABLE KEYS */;
/*!40000 ALTER TABLE `tbl_systemlog` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `tbl_task`
--

DROP TABLE IF EXISTS `tbl_task`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `tbl_task` (
  `id` int(11) NOT NULL AUTO_INCREMENT,
  `Task` longtext,
  `Status` varchar(45) DEFAULT NULL,
  `Task_Date` datetime DEFAULT NULL,
  PRIMARY KEY (`id`)
) ENGINE=InnoDB AUTO_INCREMENT=25 DEFAULT CHARSET=latin1;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `tbl_task`
--

LOCK TABLES `tbl_task` WRITE;
/*!40000 ALTER TABLE `tbl_task` DISABLE KEYS */;
INSERT INTO `tbl_task` VALUES (1,'nene','Completed','2023-10-08 00:00:00'),(2,'nene','Completed','2023-10-11 14:51:36'),(3,'nene','Assigned','2023-10-14 00:00:00'),(4,'nene','Assigned','2023-10-14 00:00:00'),(5,'nene','Assigned','2023-10-14 00:00:00'),(6,'monday','Assigned','2023-10-14 00:00:00'),(7,'monday','Assigned','2023-10-14 00:00:00'),(8,'okkkk','Assigned','2023-10-14 00:00:00'),(9,'rizal','Completed','2023-10-08 00:00:00'),(10,'rizal','Completed','2023-10-14 00:00:00'),(11,'delivery','Assigned','2023-10-14 00:00:00'),(12,'today task','Assigned','2023-10-14 00:00:00'),(16,'rizal','Assigned','2023-10-14 00:00:00'),(17,'rizal','Assigned','2023-10-14 00:00:00'),(18,'a','Assigned','2023-11-12 00:00:00'),(19,'b','Assigned','2023-11-11 00:00:00'),(20,'c','Assigned','2023-11-11 00:00:00'),(21,'d','Assigned','2023-11-12 00:00:00'),(22,'e','Assigned','2023-11-18 00:00:00'),(23,'j','Completed','2023-11-30 00:00:00'),(24,'rizal','Completed','2023-11-11 00:00:00');
/*!40000 ALTER TABLE `tbl_task` ENABLE KEYS */;
UNLOCK TABLES;
/*!40103 SET TIME_ZONE=@OLD_TIME_ZONE */;

/*!40101 SET SQL_MODE=@OLD_SQL_MODE */;
/*!40014 SET FOREIGN_KEY_CHECKS=@OLD_FOREIGN_KEY_CHECKS */;
/*!40014 SET UNIQUE_CHECKS=@OLD_UNIQUE_CHECKS */;
/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
/*!40111 SET SQL_NOTES=@OLD_SQL_NOTES */;

-- Dump completed on 2023-11-11 17:00:41
