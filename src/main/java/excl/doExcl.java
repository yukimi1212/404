package excl;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.OutputStream;
import java.io.Reader;
import java.io.UnsupportedEncodingException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.metadata.Sheet;
import com.alibaba.excel.metadata.Table;
import com.alibaba.excel.support.ExcelTypeEnum;

import excl.order.Order;

public class doExcl {

	public static void main(String[] args) throws FileNotFoundException, UnsupportedEncodingException {
		ArrayList<Order> listOrder = new ArrayList<Order>();
		listOrder = readFile();
		if(listOrder.size() != 0)
			writeFile(listOrder);
	}

	public static void writeFile(ArrayList<Order> listOrder) throws FileNotFoundException {
		// 生成EXCEL并指定输出路径
        OutputStream out = new FileOutputStream("C:\\Users\\milly\\Desktop\\孙东杓饭制证件照2.0.xlsx");
        ExcelWriter writer = new ExcelWriter(out, ExcelTypeEnum.XLSX);
 
        // 设置SHEET
        Sheet sheet = new Sheet(1, 0);
        sheet.setSheetName("sheet1");
 
        // 设置标题
        Table table = new Table(1);
        List<List<String>> titles = new ArrayList<List<String>>();
 
        titles.add(Arrays.asList("数量"));
        titles.add(Arrays.asList("时间"));
        titles.add(Arrays.asList("订单编号"));
        titles.add(Arrays.asList("商品（定金）"));
        titles.add(Arrays.asList("收件人"));
        titles.add(Arrays.asList("电话"));
        titles.add(Arrays.asList("地址"));

        table.setHead(titles);

        
        Integer writeCount = listOrder.size();
        List<List<String>> userList = new ArrayList<List<String>>();
        List<String> listString;
        for (int j = 0; j < writeCount; j++) {
        	listString = Arrays.asList(listOrder.get(j).getNum() + "", listOrder.get(j).getOrderTime(), listOrder.get(j).getOrderId(), listOrder.get(j).getOrderName(), listOrder.get(j).getBuyerName(), listOrder.get(j).getBuyerPhone(), listOrder.get(j).getBuyerAddress()); 
        	System.out.println(listString);
        	userList.add(listString);          
        }
        writer.write0(userList, sheet, table); 	
        
        writer.finish();
        System.out.println(" ***success!*** ");
	}
	
	public static ArrayList<Order> readFile() throws UnsupportedEncodingException {
        String pathname = "C:\\Users\\milly\\Desktop\\txt\\孙东杓饭制证件照2.0.txt";
          
        File  file = new File(pathname);
        String fileName = file.getName();
        String finalName = fileName.substring(0,fileName.length()-4);
        System.out.println(finalName);
        Reader reader = null;
        StringBuffer buffer = new StringBuffer();
        Order order = new Order();
        ArrayList<Order> listOrder = new ArrayList<Order>();
        boolean isStart = false;
        boolean isStartCotinue = false;
        boolean isPhoneStart = false;
        boolean isPhoneEnd = false;
        boolean isAddress = false;
        boolean isId = false;
        boolean isTime = false;
        boolean isEnd1 = false;
        boolean isEnd2 = false;
        try {

            reader = new InputStreamReader(new FileInputStream(file), "gbk");
            int temp;
            char tempchar;
            
            try {
                while((temp = reader.read())!=-1) {
                	tempchar = (char)temp;
                
                    if(tempchar != '\r') {
                    	if(!Character.isDigit(tempchar) && isStart && !isStartCotinue && !isPhoneStart) {
//                    		order.setBuyerName(order.getBuyerName() + buffer.toString());
//                    		buffer = new StringBuffer();
                    		isStartCotinue = true;
                    	}
                    	
                    	if(tempchar == ' ' && !isStart) {    		
                    		order.setBuyerName(buffer.toString());
                    		buffer = new StringBuffer();
                    		isStart = true;
                    	}
                    	
                    	if(tempchar == ' ' && isStartCotinue && isStart) {
                    		order.setBuyerName(order.getBuyerName() + " " + buffer.toString());
                    		buffer = new StringBuffer();
                    		isStartCotinue = false;
                    	}
                    	if(Character.isDigit(tempchar) && isStart && !isStartCotinue && !isPhoneStart) {
                    		buffer = new StringBuffer();
                    		isPhoneStart = true;
                    	}
                    	if(!Character.isDigit(tempchar) && isPhoneStart && !isPhoneEnd) {
                    		order.setBuyerPhone(buffer.toString());
                    		buffer = new StringBuffer();
                    		isPhoneEnd = true;
                    	}
                    	if(tempchar=='订' && isPhoneEnd && !isAddress) {
                    		order.setBuyerAddress(buffer.toString());
                    		buffer = new StringBuffer();
                    		isAddress = true;
                    	}
                    	if(tempchar=='下' && isAddress && !isId) {
                    		String id = buffer.toString().substring(5, buffer.length());
                    		order.setOrderId(id);
                    		buffer = new StringBuffer();
                    		isId = true;
                    	}
                    	if(tempchar=='商' && isId && !isTime) {
                    		String time = buffer.toString().substring(5, 15);
                    		order.setOrderTime(time);
                    		buffer = new StringBuffer();
                    		isTime = true;
                    	}
                    	if(tempchar=='订' && isTime && !isEnd1) {
                    		isEnd1 = true;
                    		continue;
                    	}
                    	if(tempchar=='备' && isEnd1) {
                    		isEnd1 = false;
                    		continue;
                    	}
                    	if(tempchar=='.' && isEnd1 && !isEnd2) {
                    		isEnd2 = true;
                    		order.setNum(listOrder.size() + 1);
                    		order.setOrderName(finalName);
                            listOrder.add(order);
                    		continue;
                    	}
                    	if(!Character.isDigit(tempchar) && tempchar != '.' && isEnd2) {
                    		isStart = false;
                    		isStartCotinue = false;
                    		isPhoneStart = false;
                            isPhoneEnd = false;
                            isAddress = false;
                            isId = false;
                            isTime = false;
                            isEnd1 = false;
                            isEnd2 = false;
                            buffer = new StringBuffer();
                            
                            order = new Order();
                    	}
             
                    	if((tempchar != '\n' && tempchar != ' ') || isStartCotinue) {
                    		buffer.append(tempchar);
                    	}
                    }
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
            try {
                reader.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }catch(FileNotFoundException e) {
            e.printStackTrace();
        }      
		return listOrder;
    }

}
