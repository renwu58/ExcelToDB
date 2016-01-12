/**
 * zhangrz3
 * 2015�?�?�?
 */
package com.jeffy.importexcel;

import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Set;

import org.apache.log4j.Logger;

/**
 * 处理数据库增删改查工具类
 * @author zhangrz3
 *
 */
public class DBTools {
	private static Logger logger = Logger.getLogger(DBTools.class);
	
	private DatabaseInfo info;
	private Connection conn;

	/**
	 * @param args
	 */
	public static void main(String[] args) {
		
		String sql="select * from ods_test where id=? and name=?";
		List<Object> conList = new ArrayList<Object>();
		conList.add(3);
		conList.add("wangwu");
		DatabaseInfo info = new DatabaseInfo();
		DBTools db = new DBTools(info);
		List<Map<String, Object>> testCollections=db.getSelect(sql, conList);
		
		for (Map<String, Object> map : testCollections) {
			Set<Entry<String, Object>> sets=map.entrySet();
			for (Entry<String, Object> entry : sets) {
				System.out.println(entry.getKey() + " ==> " + entry.getValue());
			}
		}
		
		
		//Map<String, Object> testMap = DBTools.getFirst(sql, null);
		/*
		Set<Entry<String, Object>> sets=testMap.entrySet();
		for (Entry<String, Object> entry : sets) {
			System.out.println(entry.getKey() + " ==> " + entry.getValue());
		}
		*/
	}
	public DBTools(DatabaseInfo info){
		this.info = info;
		this.conn = ConnectionFactory.getConnection(info.getUrl(), info.getUser(), info.getPassword());
	}
	/**
	 * 根据指定的查询语句返回第�?��匹配的�?
	 * @param sql 查询语句
	 * @param conditions  查询条件
	 * @return
	 */
	public Map<String, Object> getFirst(String sql,List<Object> conditions){
		Map<String, Object> resultMap = new HashMap<String, Object>();
		try {
			PreparedStatement ps = conn.prepareStatement(sql);
			ResultSet sets = null;
			if (conditions != null && conditions.size()!=0) {
				for(int i=0;i<conditions.size();i++){
					ps.setObject((i+1), conditions.get(i));
				}
			}
			if(ps.execute()){
				sets = ps.getResultSet();
				ResultSetMetaData metaData = sets.getMetaData();
				Integer len = metaData.getColumnCount();
				while (sets.next()) {
					for(int j=1;j<=len;j++){
						Object obj=sets.getObject(j);
						resultMap.put(metaData.getColumnName(j), obj);
					}
					break;
				}
			}
			sets.close();
			ps.close();
		} catch (SQLException e) {
			logger.error("Execute select failure! sql： "+ sql, e);
		}
		return resultMap;
	}
	/**
	 * 根据指定的条件获取所有满足条件的记录
	 * @param sql 查询语句
	 * @param conditions 查询条件
	 * @return List<Map<String,Object>> 
	 */
	public List<Map<String, Object>> getSelect(String sql,List<Object> conditions){
		//Data Collections
		List<Map<String, Object>> collections = new ArrayList<Map<String,Object>>();
		//Data Row
		Map<String, Object> row = new HashMap<String, Object>();
		//Get a DataSource Connection
		PreparedStatement ps = null;
		ResultSet sets = null;
		
		try {
			ps=conn.prepareStatement(sql);
			//if have condition,set the condition.
			if (conditions != null && conditions.size()!=0) {
				for(int i=0;i<conditions.size();i++){
					ps.setObject((i+1),conditions.get(i));
				}
			}
			if (ps.execute()) {
				sets = ps.getResultSet();
				ResultSetMetaData metaData = sets.getMetaData();
				Integer len = metaData.getColumnCount();
				while (sets.next()) {
					for(int j=1;j<=len;j++){
						Object obj=sets.getObject(j);
						row.put(metaData.getColumnName(j), obj);
					}
					collections.add(row);
				}
			}
			
		} catch (SQLException e) {
			logger.error("Execute select failure! sql： "+ sql, e);
		}finally{
			try {
				sets.close();
				ps.close();
			} catch (SQLException e) {
				logger.error("Can not close resource!", e);
			}
		}
		return collections;
	}
	/**
	 * 进行DML操作的统�?���?
	 * @param sql 要执行的SQL脚本
	 * @param conditions 执行的条�?
	 * @return Integer 受影响的行数
	 */
	public Integer doUpdate(String sql,List<Object> conditions){
		Integer affects = 0;
		PreparedStatement ps = null;
		try {
			ps=conn.prepareStatement(sql);
			if (conditions != null && conditions.size()!=0) {
				for(int i=0;i<conditions.size();i++){
					ps.setObject((i+1),conditions.get(i));
				}
			}
			affects=ps.executeUpdate();
		} catch (SQLException e) {
			logger.error("Update failure! sql： "+sql+" condition： "+conditions.toString(), e);
		}finally{
			try {
				ps.close();
			} catch (SQLException e) {
				logger.error("Can not close resource!", e);
			}
		}
		return affects;
	}
	public Integer doCreate(String sql){
		Statement state = null;
		Integer affects = 0;
		try {
			state = conn.createStatement();
			state.execute(sql);
			affects = state.getUpdateCount();
		} catch (SQLException e) {
			logger.error("Execute sql: " + sql +" failure!" , e);
		}finally{
			try {
				state.close();
			} catch (SQLException e) {
				logger.error("Can not close resource!", e);
			}
		}
		return affects;
	}
	public void close(){
		try {
			this.conn.close();
		} catch (SQLException e) {
			logger.error("Can not close connection resource!", e);
		}
	}
}
