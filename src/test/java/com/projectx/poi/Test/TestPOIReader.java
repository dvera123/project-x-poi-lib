package com.projectx.poi.Test;

import java.util.List;

import org.junit.Test;
import com.projectx.poi.util.*;

public class TestPOIReader {
	
	@Test
	public void poiReaderTest(){
		
		POIReader reader = new POIReader();
		String fileName = "test.xls";
		List<Object> l = reader.getDataPerSheet("", fileName, 0);
		
		for(Object row : l){
			System.out.println(row.toString());
		}
	}

}
