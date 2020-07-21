package com.lan.excel.main.report;

import java.util.ArrayList;
import java.util.List;

public class ExampleData {
	
	public List<List<Object>> getData1(int num){
		List<List<Object>> list=new ArrayList<List<Object>>(); 
		for(int i=0; i<num; i++){
			List<Object> list1 = new ArrayList<>();
			list1.add("产品-"+i);
			list1.add((int)(Math.random()*200));
			list1.add(Math.random()*100000);
			list.add(list1);
		}
		return list;
	}
	
	public List<List<Object>> getData2(){
		List<List<Object>> list=new ArrayList<List<Object>>();
		String[] provinces=new String[]{"湖南", "湖北", "江苏", "福建"};
		for(int i=0; i<4; i++){
			List<Object> list1 = new ArrayList<>();
			list1.add(i+1);
			list1.add(provinces[i]);
			list1.add("医用防护口罩");
			list1.add((int)(Math.random()*10000000));
			double income=Math.random()*20000000;
			list1.add(income);
			list1.add(income*0.65);
			list.add(list1);
		}
		return list;
	}
	
	public List<List<Object>> getData3(){
		List<List<Object>> list=new ArrayList<List<Object>>();
		String[] months=new String[]{"2019-08", "2019-09", "2019-10", "2019-11", "2019-12",
				"2020-01", "2020-02", "2020-03", "2020-04", "2020-05", "2020-06", "2020-07"};
		for(int i=0; i<12; i++){
			List<Object> list1 = new ArrayList<>();
			list1.add(months[i]);
			list1.add((int)(Math.random()*2000));
			list1.add((int)(Math.random()*1900));
			list.add(list1);
		}
		return list;
	}

}
