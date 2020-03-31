package com.forcewin.excelWriter.service;

import java.util.HashMap;
import java.util.Map;

public class Service {
    public static String getConnectionURL(){
        return "jdbc:mariadb://127.0.0.1:3306/?";
    }
    public static String getClassName(){
        return "org.mariadb.jdbc.Driver";
    }
    public static String getRepoID(){
        return "guest";
    }
    public static String getRepoPW(){
        return "guest123";
    }
    public static String getRepoName(){
        return "mariadb";
    }
    public static boolean isSpecialRepo(String repo_name) {
		repo_name = repo_name.toLowerCase();
		if (repo_name.equals("oracle")) {
			return true;
		} else if (repo_name.equals("mssql")) {
			return true;
		} else if (repo_name.equals("tibero")) {
			return true;
		} else {
			return false;
		}
    }
    public static Map<String,Boolean> getDBFilter(){
        Map <String,Boolean> resultMap = new HashMap<String,Boolean>();
        resultMap.put("information_schema",true);
        resultMap.put("mysql",true);
        resultMap.put("performance_schema",true);
        return resultMap;
    }

    public static Map<String,Boolean> getTableFilter(){
        Map <String,Boolean> resultMap = new HashMap<String,Boolean>();
        return resultMap;
    }
    public static String getViewoption(){
        return "skip";
    }
    public static Map<String,String> getMetaDataMapper(){
        Map <String,String> resultMap = new HashMap<String,String>();
        resultMap.put("db1,table1_1,FLD1_1_1","SHA256");
        resultMap.put("db2,table2_2,FLD2_2_3","HASH");
        return resultMap;
    }
}