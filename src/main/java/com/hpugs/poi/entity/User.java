package com.hpugs.poi.entity;

import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public class User {
	
	private Integer id;
	private String realName;
	private String nickName;
	private Integer age;
	private double monery;
	private Date gmtCreat;
	private String	remark;
	
	public Integer getId() {
		return id;
	}
	public void setId(Integer id) {
		this.id = id;
	}
	public String getRealName() {
		return realName;
	}
	public void setRealName(String realName) {
		this.realName = realName;
	}
	public String getNickName() {
		return nickName;
	}
	public void setNickName(String nickName) {
		this.nickName = nickName;
	}
	public Integer getAge() {
		return age;
	}
	public void setAge(Integer age) {
		this.age = age;
	}
	public double getMonery() {
		return monery;
	}
	public void setMonery(double monery) {
		this.monery = monery;
	}
	public Date getGmtCreat() {
		return gmtCreat;
	}
	public void setGmtCreat(Date gmtCreat) {
		this.gmtCreat = gmtCreat;
	}
	public String getRemark() {
		return remark;
	}
	public void setRemark(String remark) {
		this.remark = remark;
	}
	
	@Override
	public String toString() {
		return "User [id=" + id + ", realName=" + realName + ", nickName=" + nickName + ", age=" + age + ", monery="
				+ monery + ", gmtCreat=" + gmtCreat + ", remark=" + remark + "]";
	}
	 
	public static List<User> createUserList(Integer userCount){
		if(null == userCount){
			userCount = 100;
		}
		List<User> users = new ArrayList<User>();
		for(int i=0; i<userCount; i++){
			User user = new User();
			user.setId(i+1);
			user.setRealName("gs"+i);
			user.setNickName("hpugs"+i);
			user.setAge(20+i);
			user.setMonery(34.23*i);
			user.setGmtCreat(new Date());
			user.setRemark("备注：\n i="+i);
			users.add(user);
		}
		return users;
	}
	
}
