-- phpMyAdmin SQL Dump
-- version 4.6.6
-- https://www.phpmyadmin.net/
--
-- Servidor: localhost
-- Tiempo de generación: 24-01-2023 a las 08:14:03
-- Versión del servidor: 5.7.17-log
-- Versión de PHP: 5.6.30

SET SQL_MODE = "NO_AUTO_VALUE_ON_ZERO";
SET time_zone = "+00:00";


/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8mb4 */;

--
-- Base de datos: `ci`
--

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `log_ci`
--

CREATE TABLE `log_ci` (
  `codlog` bigint(20) NOT NULL,
  `fecha` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  `cadena` varchar(255) NOT NULL,
  `in_b5` tinyint(4) NOT NULL,
  `in_b10` tinyint(4) NOT NULL,
  `in_b20` tinyint(4) NOT NULL,
  `in_b50` tinyint(4) NOT NULL,
  `in_b100` tinyint(4) NOT NULL,
  `in_b200` tinyint(4) NOT NULL,
  `out_b5` tinyint(4) NOT NULL,
  `out_b10` tinyint(4) NOT NULL,
  `out_b20` tinyint(4) NOT NULL,
  `out_b50` tinyint(4) NOT NULL,
  `out_b100` tinyint(4) NOT NULL,
  `out_b200` tinyint(4) NOT NULL,
  `in_m5c` tinyint(4) NOT NULL,
  `in_m10c` tinyint(4) NOT NULL,
  `in_m20c` tinyint(4) NOT NULL,
  `in_m50c` tinyint(4) NOT NULL,
  `in_m1` tinyint(4) NOT NULL,
  `in_m2` tinyint(4) NOT NULL,
  `out_m5c` tinyint(4) NOT NULL,
  `out_m10c` tinyint(4) NOT NULL,
  `out_m20c` tinyint(4) NOT NULL,
  `out_m50c` tinyint(4) NOT NULL,
  `out_m1` tinyint(4) NOT NULL,
  `out_m2` tinyint(4) NOT NULL,
  `lim_m5c` smallint(6) NOT NULL,
  `lim_m10c` smallint(6) NOT NULL,
  `lim_m20c` smallint(6) NOT NULL,
  `lim_m50c` smallint(6) NOT NULL,
  `lim_m1` smallint(6) NOT NULL,
  `lim_m2` smallint(6) NOT NULL,
  `lim_b5` smallint(6) NOT NULL,
  `lim_b10` smallint(6) NOT NULL,
  `lim_b20` smallint(6) NOT NULL,
  `lim_b50` smallint(6) NOT NULL,
  `lim_b100` smallint(6) NOT NULL,
  `niv_m5c` smallint(6) NOT NULL,
  `niv_m10c` smallint(6) NOT NULL,
  `niv_m20c` smallint(6) NOT NULL,
  `niv_m50c` smallint(6) NOT NULL,
  `niv_m1` smallint(6) NOT NULL,
  `niv_m2` smallint(6) NOT NULL,
  `niv_b5` smallint(6) NOT NULL,
  `niv_b10` smallint(6) NOT NULL,
  `niv_b20` smallint(6) NOT NULL,
  `niv_b50` smallint(6) NOT NULL,
  `niv_b100` smallint(6) NOT NULL
) ENGINE=MyISAM DEFAULT CHARSET=utf8;

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `parametros`
--

CREATE TABLE `parametros` (
  `posicion` tinyint(4) NOT NULL,
  `valor` float NOT NULL,
  `comportamiento` char(1) NOT NULL,
  `max` smallint(6) NOT NULL,
  `min` smallint(6) NOT NULL,
  `actual` int(11) NOT NULL
) ENGINE=MyISAM DEFAULT CHARSET=utf8mb4;

--
-- Volcado de datos para la tabla `parametros`
--

INSERT INTO `parametros` (`posicion`, `valor`, `comportamiento`, `max`, `min`, `actual`) VALUES
(0, 0.01, '-', 30, 20, 16),
(1, 0.02, '-', 30, 20, 18),
(2, 0.05, '-', 30, 20, 11),
(3, 0.1, '-', 30, 20, 10),
(4, 0.2, '-', 30, 20, 15),
(5, 0.5, '-', 30, 20, 12),
(6, 1, '-', 30, 20, 22),
(7, 2, '-', 30, 20, 25),
(8, 5, '-', 30, 20, 5),
(9, 10, '-', 30, 20, 1),
(10, 20, '-', 30, 10, 15),
(11, 50, '0', 0, 0, 1),
(12, 100, '0', 0, 0, 0),
(13, 200, '3', 0, 0, 0),
(14, 500, '3', 0, 0, 0);

--
-- Índices para tablas volcadas
--

--
-- Indices de la tabla `log_ci`
--
ALTER TABLE `log_ci`
  ADD PRIMARY KEY (`codlog`);

--
-- Indices de la tabla `parametros`
--
ALTER TABLE `parametros`
  ADD PRIMARY KEY (`posicion`);

--
-- AUTO_INCREMENT de las tablas volcadas
--

--
-- AUTO_INCREMENT de la tabla `log_ci`
--
ALTER TABLE `log_ci`
  MODIFY `codlog` bigint(20) NOT NULL AUTO_INCREMENT;
/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
