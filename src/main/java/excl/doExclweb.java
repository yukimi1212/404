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

public class doExclweb {
	
	public static String fileName = "裴珠泫 Carrot Irene🐇🥕";
	
	public static void main(String[] args) throws FileNotFoundException, UnsupportedEncodingException {
		ArrayList<Order> listOrder = new ArrayList<Order>();
		listOrder = readFile();
		if(listOrder.size() != 0)
			writeFile(listOrder);
	}

	public static void writeFile(ArrayList<Order> listOrder) throws FileNotFoundException {	
        OutputStream out = new FileOutputStream("C:\\Users\\milly\\Desktop\\txt\\"+ fileName + ".xlsx");
        ExcelWriter writer = new ExcelWriter(out, ExcelTypeEnum.XLSX);	// 生成EXCEL并指定输出路径
        
        Sheet sheet = new Sheet(1, 0);	
        sheet.setSheetName("sheet1");	// 设置SHEET      
        Table table = new Table(1);		// 设置标题
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
        String pathname = "C:\\Users\\milly\\Desktop\\txt\\定金表格信息复制.txt";
          
        File  file = new File(pathname);
        System.out.println(fileName);
        Reader reader = null;
        StringBuffer buffer = new StringBuffer();
        Order order = new Order();
        ArrayList<Order> listOrder = new ArrayList<Order>();
        boolean isStart = false;
        boolean isStartColon = false;
        boolean isPhone = false;
        boolean isAddress = false;
        boolean isAddress_Xiu = false;
        boolean isAddress_Di = false;
        boolean isEndAddress = false;
        boolean isNearId = false;
        boolean isId = false;
        boolean isEndId = false;
        boolean isTime = false;
        boolean isEnd = false;

        try {

            reader = new InputStreamReader(new FileInputStream(file), "gbk");
            int temp;
            char tempchar;
            
            try {
                while((temp = reader.read())!=-1) {
                	tempchar = (char)temp;
                
                    if(tempchar != '\r') {
                    	if (tempchar=='址' && !isStart) {
                    		isStart = true;
                    	}
                    	
                    	if (tempchar=='：' && !isStartColon && isStart) {
                    		buffer = new StringBuffer();
                    		isStartColon = true;
                    	}
                    	
                    	if (tempchar=='，' && isStartColon && !isPhone) {
                    		order.setBuyerName(buffer.toString().substring(1));
                    		buffer = new StringBuffer();
                    		isPhone = true;
                    	}
                    	                   	
                    	if (tempchar=='，' && buffer.length() != 0 && isPhone && !isAddress) {
                    		order.setBuyerPhone(buffer.toString());
                    		buffer = new StringBuffer();
                    		isAddress = true;
                    	}
                    	
                    	if (tempchar=='修' && isAddress && !isEndAddress) {
                    		isAddress_Xiu = true;
                    	}
                    	
                    	if (tempchar=='改' && isAddress && isAddress_Xiu && !isEndAddress) {
                    		order.setBuyerAddress(buffer.toString().substring(0,buffer.toString().length()-1));
                    		buffer = new StringBuffer();
                    		isEndAddress = true;
                    	}
                    	
                    	if (tempchar=='地' && isAddress && !isEndAddress) {
                    		isAddress_Di = true;
                    	}
                    	
                    	if (tempchar=='址' && isAddress && isAddress_Di && !isEndAddress) {
                    		order.setBuyerAddress(buffer.toString().substring(0,buffer.toString().length()-1));
                    		buffer = new StringBuffer();
                    		isEndAddress = true;
                    	}
                    	
                    	if (tempchar=='账' && isAddress && !isEndAddress) {
                    		order.setBuyerAddress(buffer.toString().substring(0,buffer.toString().length()-6));
                    		buffer = new StringBuffer();
                    		isEndAddress = true;
                    	}
                    	
                    	if (tempchar=='编' && isEndAddress && !isNearId) {
                    		isNearId = true;
                    	}
                    	
                    	if (tempchar=='：' && isNearId && !isId) {
                    		buffer = new StringBuffer();
                    		isId = true;
                    	}
                    	
                    	if (tempchar=='下' && isId && !isEndId) {
                    		order.setOrderId(buffer.toString().substring(1));
                    		buffer = new StringBuffer();
                    		isEndId = true;
                    	}
                    	
                    	if (tempchar=='：' && isEndId && !isTime) {
                    		buffer = new StringBuffer();
                    		isTime = true;
                    	}
                    	
                    	if (tempchar==':' && buffer.length() != 0 && isTime && !isEnd) {
                    		String time = buffer.toString().substring(1, 11);
                    		order.setOrderTime(time);
                    		isEnd = true;
                    	}
             
                    	if(isEnd) {
                    		order.setNum(listOrder.size() + 1);
                    		order.setOrderName(fileName);
                            listOrder.add(order);
                            
                            isStart = false;
                            isStartColon = false;
                            isPhone = false;
                            isAddress = false;
                            isAddress_Xiu = false;
                            isAddress_Di = false;
                            isEndAddress = false;
                            isNearId = false;
                            isId = false;
                            isEndId = false;
                            isTime = false;
                            isEnd = false;
                            
                            buffer = new StringBuffer();     
                            order = new Order();
                    	}
              
                    	if ((tempchar != '\n' && tempchar != ' ' && tempchar != '，')) {
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
