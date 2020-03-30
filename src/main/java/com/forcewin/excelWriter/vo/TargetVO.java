package com.forcewin.excelWriter.vo;

public class TargetVO{
    private int target_id;
    private String target_name;
    private int repo_cd;
    private String repo_user_id;
    private String repo_user_pw;
    private String repo_ip;
    private int repo_port;

    public int getTarget_id() {
        return target_id;
    }

    public void setTarget_id(int target_id) {
        this.target_id = target_id;
    }

    public String getTarget_name() {
        return target_name;
    }

    public void setTarget_name(String target_name) {
        this.target_name = target_name;
    }

    public int getRepo_cd() {
        return repo_cd;
    }

    public void setRepo_cd(int repo_cd) {
        this.repo_cd = repo_cd;
    }

    public String getRepo_user_id() {
        return repo_user_id;
    }

    public void setRepo_user_id(String repo_user_id) {
        this.repo_user_id = repo_user_id;
    }

    public String getRepo_user_pw() {
        return repo_user_pw;
    }

    public void setRepo_user_pw(String repo_user_pw) {
        this.repo_user_pw = repo_user_pw;
    }

    public String getRepo_ip() {
        return repo_ip;
    }

    public void setRepo_ip(String repo_ip) {
        this.repo_ip = repo_ip;
    }

    public int getRepo_port() {
        return repo_port;
    }

    public void setRepo_port(int repo_port) {
        this.repo_port = repo_port;
    }
    
}