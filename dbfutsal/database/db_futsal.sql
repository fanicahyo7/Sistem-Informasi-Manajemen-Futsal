-- phpMyAdmin SQL Dump
-- version 4.1.6
-- http://www.phpmyadmin.net
--
-- Host: 127.0.0.1
-- Generation Time: Jul 26, 2016 at 12:24 PM
-- Server version: 5.6.16
-- PHP Version: 5.5.9

SET SQL_MODE = "NO_AUTO_VALUE_ON_ZERO";
SET time_zone = "+00:00";


/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8 */;

--
-- Database: `db_futsal`
--

-- --------------------------------------------------------

--
-- Table structure for table `accesslevel`
--

CREATE TABLE IF NOT EXISTS `accesslevel` (
  `LevelID` varchar(100) NOT NULL,
  `LevelName` varchar(100) NOT NULL,
  PRIMARY KEY (`LevelID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `accesslevel`
--

INSERT INTO `accesslevel` (`LevelID`, `LevelName`) VALUES
('1', 'Admin'),
('2', 'Kasir');

-- --------------------------------------------------------

--
-- Table structure for table `itempembelian`
--

CREATE TABLE IF NOT EXISTS `itempembelian` (
  `NoUrut` int(50) NOT NULL,
  `KodePembelian` varchar(100) NOT NULL,
  `Tanggal` date NOT NULL,
  `KodeBarang` varchar(100) NOT NULL,
  `Jumlah` int(100) NOT NULL,
  `Harga` int(100) NOT NULL,
  `Total` int(100) NOT NULL,
  PRIMARY KEY (`NoUrut`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `itempembelian`
--

INSERT INTO `itempembelian` (`NoUrut`, `KodePembelian`, `Tanggal`, `KodeBarang`, `Jumlah`, `Harga`, `Total`) VALUES
(5, 'PBL004', '2014-11-05', 'BRG001', 5, 100000, 500000);

-- --------------------------------------------------------

--
-- Table structure for table `itempenjualan`
--

CREATE TABLE IF NOT EXISTS `itempenjualan` (
  `NoUrut` int(50) NOT NULL,
  `KodePenjualan` varchar(100) NOT NULL,
  `KodeBarang` varchar(100) NOT NULL,
  `Tanggal` date NOT NULL,
  `NoPakaiLapangan` varchar(100) NOT NULL,
  `Jumlah` int(100) NOT NULL,
  `Harga` int(100) NOT NULL,
  `Total` int(100) NOT NULL,
  PRIMARY KEY (`NoUrut`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `itempenjualan`
--

INSERT INTO `itempenjualan` (`NoUrut`, `KodePenjualan`, `KodeBarang`, `Tanggal`, `NoPakaiLapangan`, `Jumlah`, `Harga`, `Total`) VALUES
(30, 'JUL0511144001', 'BRG001', '2014-11-05', 'PKL0511140001', 1, 160000, 160000),
(31, 'JUL2607160001', 'BRG001', '2016-07-26', 'PKL2607160001', 1, 160000, 160000);

-- --------------------------------------------------------

--
-- Table structure for table `jenismember`
--

CREATE TABLE IF NOT EXISTS `jenismember` (
  `KodeJenisMember` varchar(100) NOT NULL,
  `NamaJenisMember` varchar(100) NOT NULL,
  `JumlahBulan` int(50) NOT NULL,
  `Biaya` int(100) NOT NULL,
  PRIMARY KEY (`KodeJenisMember`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `jenismember`
--

INSERT INTO `jenismember` (`KodeJenisMember`, `NamaJenisMember`, `JumlahBulan`, `Biaya`) VALUES
('KMB002', 'Selamanya', 0, 200000),
('KMB001', '1 Bulan', 1, 50000);

-- --------------------------------------------------------

--
-- Table structure for table `login`
--

CREATE TABLE IF NOT EXISTS `login` (
  `KodeUser` varchar(100) NOT NULL,
  `UserName` varchar(100) NOT NULL,
  `UserPsw` varchar(100) NOT NULL,
  `Alamat` text NOT NULL,
  `HPNo` int(100) NOT NULL,
  `LevelID` int(100) NOT NULL,
  PRIMARY KEY (`UserName`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `login`
--

INSERT INTO `login` (`KodeUser`, `UserName`, `UserPsw`, `Alamat`, `HPNo`, `LevelID`) VALUES
('USR001', 'Admin', 'ADMIN', 'jl.akjgdsfhbswgfvrwg', 2147483647, 1),
('', 'Fani', 'Dwi', 'Jl.Mojorejo', 898676788, 2),
('USR002', 'Kasir', 'Kasir', 'Jl.Anu', 93467, 2),
('USR003', 'Cahyo', 'cahyo', 'Jl.batu', 2147483647, 2);

-- --------------------------------------------------------

--
-- Table structure for table `member`
--

CREATE TABLE IF NOT EXISTS `member` (
  `NoRegister` varchar(100) NOT NULL,
  `NamaTim` varchar(100) NOT NULL,
  `AtasNama` varchar(100) NOT NULL,
  `Alamat` text NOT NULL,
  `Telp` int(100) NOT NULL,
  `isMember` int(50) NOT NULL,
  `TanggalDaftar` date NOT NULL,
  `TanggalHangus` date NOT NULL,
  `Biaya` int(50) NOT NULL,
  `Berbatas` int(50) NOT NULL,
  `KodeJenisMember` varchar(100) NOT NULL,
  PRIMARY KEY (`NoRegister`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `member`
--

INSERT INTO `member` (`NoRegister`, `NamaTim`, `AtasNama`, `Alamat`, `Telp`, `isMember`, `TanggalDaftar`, `TanggalHangus`, `Biaya`, `Berbatas`, `KodeJenisMember`) VALUES
('MMB001', 'Kali Putih F.C', 'Dani', 'Jl.Ketindan', 838395854, 0, '2014-11-05', '2014-12-05', 50000, 0, 'KMB001'),
('MMB002', 'Sparta', 'Teguh', 'Jl.Mulyoagung', 2147483647, -1, '2016-07-26', '2016-08-26', 50000, 0, 'KMB001');

-- --------------------------------------------------------

--
-- Table structure for table `mstbarang`
--

CREATE TABLE IF NOT EXISTS `mstbarang` (
  `KodeBarang` varchar(100) NOT NULL,
  `NamaBarang` varchar(100) NOT NULL,
  `HargaJual` int(50) NOT NULL,
  `HargaBeli` int(50) NOT NULL,
  `Stok` int(50) NOT NULL,
  `Foto` longblob NOT NULL,
  `Lokasi` varchar(100) NOT NULL,
  PRIMARY KEY (`KodeBarang`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `mstbarang`
--

INSERT INTO `mstbarang` (`KodeBarang`, `NamaBarang`, `HargaJual`, `HargaBeli`, `Stok`, `Foto`, `Lokasi`) VALUES
('BRG001', 'Jersey Dortmund', 160000, 100000, 2, 0x671205d5, '');

-- --------------------------------------------------------

--
-- Table structure for table `mstlapangan`
--

CREATE TABLE IF NOT EXISTS `mstlapangan` (
  `KodeLapangan` varchar(100) NOT NULL,
  `NamaLapangan` varchar(100) NOT NULL,
  `Foto` longblob NOT NULL,
  `Lokasi` varchar(500) NOT NULL,
  PRIMARY KEY (`KodeLapangan`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `mstlapangan`
--

INSERT INTO `mstlapangan` (`KodeLapangan`, `NamaLapangan`, `Foto`, `Lokasi`) VALUES
('FC001', 'Brantas', 0xbe1205d2, '');

-- --------------------------------------------------------

--
-- Table structure for table `mstshift`
--

CREATE TABLE IF NOT EXISTS `mstshift` (
  `KodeShift` varchar(100) NOT NULL,
  `JamMulai` varchar(50) NOT NULL,
  `Harga` int(50) NOT NULL,
  `HargaMember` int(50) NOT NULL,
  PRIMARY KEY (`KodeShift`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `mstshift`
--

INSERT INTO `mstshift` (`KodeShift`, `JamMulai`, `Harga`, `HargaMember`) VALUES
('SHF001', '18:00:00', 70000, 60000);

-- --------------------------------------------------------

--
-- Table structure for table `mstsupplier`
--

CREATE TABLE IF NOT EXISTS `mstsupplier` (
  `KodeSupplier` varchar(100) NOT NULL,
  `NamaSupplier` varchar(100) NOT NULL,
  `Alamat` text NOT NULL,
  `Telp` int(50) NOT NULL,
  `PenanggungJawab` varchar(100) NOT NULL,
  `HPPenanggungJawab` int(50) NOT NULL,
  PRIMARY KEY (`KodeSupplier`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `mstsupplier`
--

INSERT INTO `mstsupplier` (`KodeSupplier`, `NamaSupplier`, `Alamat`, `Telp`, `PenanggungJawab`, `HPPenanggungJawab`) VALUES
('SPL001', 'Adi Kencana', 'Jl.Subrto', 82356894, 'Hendar', 98574);

-- --------------------------------------------------------

--
-- Table structure for table `trbooking`
--

CREATE TABLE IF NOT EXISTS `trbooking` (
  `NoBooking` varchar(100) NOT NULL,
  `Tanggal` date NOT NULL,
  `TanggalBooking` date NOT NULL,
  `JamMulai` varchar(100) NOT NULL,
  `JamSelesai` varchar(100) NOT NULL,
  `KodeLapangan` varchar(100) NOT NULL,
  `DP` int(50) NOT NULL,
  `StatusBooking` int(1) NOT NULL,
  `Pembatalan` int(1) NOT NULL,
  `NoRegister` varchar(50) NOT NULL,
  `KodeShift` varchar(100) NOT NULL,
  `Harga` int(50) NOT NULL,
  `Atasnama` varchar(100) NOT NULL,
  PRIMARY KEY (`NoBooking`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `trbooking`
--

INSERT INTO `trbooking` (`NoBooking`, `Tanggal`, `TanggalBooking`, `JamMulai`, `JamSelesai`, `KodeLapangan`, `DP`, `StatusBooking`, `Pembatalan`, `NoRegister`, `KodeShift`, `Harga`, `Atasnama`) VALUES
('BOK0511140001', '2014-11-05', '2014-11-06', '18:00:00', '20:00:00', 'FC001', 80000, 0, 0, 'MMB001', 'SHF001', 120000, 'Dani'),
('BOK2607160001', '2016-07-26', '2016-07-26', '18:00:00', '19:00:00', 'FC001', 30000, 0, 0, 'mmb002', 'SHF001', 60000, 'Teguh');

-- --------------------------------------------------------

--
-- Table structure for table `trpakailapangan`
--

CREATE TABLE IF NOT EXISTS `trpakailapangan` (
  `NoPakaiLapangan` varchar(100) NOT NULL,
  `NoBooking` varchar(50) NOT NULL,
  `KodeLapangan` varchar(100) NOT NULL,
  `DP` int(50) NOT NULL,
  `HargaSewaLapangan` int(50) NOT NULL,
  `TotalPembelian` int(50) NOT NULL,
  `GrandTotalharga` int(50) NOT NULL,
  `Tanggal` date NOT NULL,
  PRIMARY KEY (`NoPakaiLapangan`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `trpakailapangan`
--

INSERT INTO `trpakailapangan` (`NoPakaiLapangan`, `NoBooking`, `KodeLapangan`, `DP`, `HargaSewaLapangan`, `TotalPembelian`, `GrandTotalharga`, `Tanggal`) VALUES
('PKL0511140001', 'BOK0511140001', 'FC001', 80000, 120000, 160000, 200000, '2014-11-05'),
('PKL2607160001', 'BOK2607160001', 'FC001', 30000, 60000, 160000, 190000, '2016-07-26');

-- --------------------------------------------------------

--
-- Table structure for table `trpembelian`
--

CREATE TABLE IF NOT EXISTS `trpembelian` (
  `KodePembelian` varchar(50) NOT NULL,
  `KodeSupplier` varchar(50) NOT NULL,
  `JumlahPembelian` int(50) NOT NULL,
  PRIMARY KEY (`KodePembelian`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `trpembelian`
--

INSERT INTO `trpembelian` (`KodePembelian`, `KodeSupplier`, `JumlahPembelian`) VALUES
('PBL004', 'SPL001', 500000);

-- --------------------------------------------------------

--
-- Table structure for table `trrevisistok`
--

CREATE TABLE IF NOT EXISTS `trrevisistok` (
  `KodeTrans` varchar(100) NOT NULL,
  `Tanggal` date NOT NULL,
  `NoUrut` int(50) NOT NULL,
  `KodeBarang` varchar(100) NOT NULL,
  `StokLama` int(50) NOT NULL,
  `StokBaru` int(50) NOT NULL,
  `Keterangan` varchar(100) NOT NULL,
  PRIMARY KEY (`KodeTrans`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `trrevisistok`
--

INSERT INTO `trrevisistok` (`KodeTrans`, `Tanggal`, `NoUrut`, `KodeBarang`, `StokLama`, `StokBaru`, `Keterangan`) VALUES
('TR001', '2014-11-05', 1, 'BRG001', 6, 4, '');

/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
