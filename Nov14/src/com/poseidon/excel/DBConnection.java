package com.poseidon.excel;

import java.sql.Connection;
import java.sql.DriverManager;

public class DBConnection {
	public static Connection getConnection() {
		Connection conn = null;
		try {
			Class.forName("org.mariadb.jdbc.Driver");
			String url = "jdbc:mariadb://localhost:3306/iot";
			conn = DriverManager.getConnection(url, "iot", "01234567");
		} catch (Exception e) {
			e.printStackTrace();
		}
		return conn;
	}

}