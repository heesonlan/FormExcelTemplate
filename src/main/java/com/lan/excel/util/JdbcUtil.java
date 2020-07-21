package com.lan.excel.util;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * 
 * @author LAN
 * @date 2020年3月19日
 */
public final class JdbcUtil {

	static Logger log = LoggerFactory.getLogger(JdbcUtil.class);
	static final String jdbcUrl = "jdbc:mysql://127.0.0.1:3306/my_db?useUnicode=true&useJDBCCompliantTimezoneShift=true&useLegacyDatetimeCode=false&serverTimezone=UTC";
	static final String userName = "root";
	static final String password = "root";
	
	private static Connection getConn(){
		return getConn0();
	}
	
	private static void releaseConn(Connection conn){
		try {conn.close();} catch (SQLException e) {e.printStackTrace();}
	}
	
	private static Connection getConn0(){
		try {
			Class.forName("com.mysql.cj.jdbc.Driver");
			Connection conn = DriverManager.getConnection(jdbcUrl, userName, password);
			return conn;
		} catch (ClassNotFoundException e) {
			log.error("未找到JDBC驱动", e);
			throw new RuntimeException("未找到JDBC驱动",e);
		} catch (SQLException e) {
			log.error("获取数据库连接失败", e);
			throw new RuntimeException("获取数据库连接失败",e);
		}
	}
	
	public static List<List<Object>> selectSqlToList(String sql, Object... param){
		sql = sql.trim();
		if(!"select".equals(sql.substring(0, 6).toLowerCase())){
			log.error("查询必须以select开始");
			throw new RuntimeException("查询必须以select开始");
		}
		Connection conn = getConn();
		try {
			PreparedStatement pstm = conn.prepareStatement(sql);
			if(param!=null){
				for (int i = 0; i < param.length; i++) {
					pstm.setObject(i + 1, param[i]);
	            }
			}
			ResultSet rs = pstm.executeQuery();
			ResultSetMetaData metaData = rs.getMetaData();
			int columnCount = metaData.getColumnCount();
			List<List<Object>> list = new ArrayList<>();
			while (rs.next()) {
				List<Object> row = new ArrayList<>();
				for(int i=1; i<=columnCount; i++){
					Object cell = rs.getObject(i);
					row.add(cell);
				}
				list.add(row);
			}
			return list;
		} catch (SQLException e) {
			log.error("查询数据库失败");
			throw new RuntimeException("查询数据库失败", e);
		} finally {
			releaseConn(conn);
		}
	}
	
	public static List<Map<String,Object>> selectSqlToMap(String sql, Object... param){
		sql = sql.trim();
		if(!"select".equals(sql.substring(0, 6).toLowerCase())){
			log.error("查询必须以select开始");
			throw new RuntimeException("查询必须以select开始");
		}
		Connection conn = getConn();
		try {
			PreparedStatement pstm = conn.prepareStatement(sql);
			if(param!=null){
				for (int i = 0; i < param.length; i++) {
					pstm.setObject(i + 1, param[i]);
	            }
			}
			ResultSet rs = pstm.executeQuery();
			ResultSetMetaData metaData = rs.getMetaData();
			int columnCount = metaData.getColumnCount();
			List<Map<String,Object>> list = new ArrayList<>();
			while (rs.next()) {
				Map<String, Object> row = new HashMap<>();
				for(int i=1; i<=columnCount; i++){
					String columnName = metaData.getColumnName(i);
					Object cell = rs.getObject(i);
					row.put(columnName, cell);
				}
				list.add(row);
			}
			return list;
		} catch (SQLException e) {
			log.error("查询数据库失败");
			throw new RuntimeException("查询数据库失败", e);
		} finally {
			releaseConn(conn);
		}
	}
	
}
