package com.forcewin.excelWriter.service;

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
}