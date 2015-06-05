package com.lydj.Controller.util;

import java.util.Map;
import java.util.Set;

/**
 * 一些系统信息，与constants类似
 * @author css001
 *
 */
public class ConInfo {


	public static Map<String,String> TOWN = null;
	
	
	
	static{
		setTown();
	}
	public static void  setTown(){
		/*TOWN = new HashMap<String,String>();
		
		TOWN.put("1", "海淀区");
		TOWN.put("2", "昌平区");
		TOWN.put("3", "西城区");*/
		TOWN = Constants.allTown;
	}
	
	public static Map<String,String>  getStreet(String id){
		Map<String,String> street = Constants.allStreet.get(id);
		return street;
	}
	public static String  getChangeStreet(String id){
		Map<String,String> street = Constants.allStreet.get(id);
		String jsons = "[";
		//  {"name": "Afghanistan", "code": "AF"}, 
		Set<String> set  = street.keySet();
		for(String s : set){
			jsons +=  "{'code': '"+s+"', 'name': '"+street.get(s)+"'}, ";
		}
		jsons = jsons.contains(",")?(jsons.substring(0,jsons.length()-2)+"]"):"";
		return jsons;
	}
}
