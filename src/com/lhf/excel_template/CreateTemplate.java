package com.lhf.excel_template;

import java.io.File;
import java.io.FileOutputStream;
import java.util.List;

import org.apache.commons.io.FileUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.DVConstraint;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFDataValidation;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.jdom.Attribute;
import org.jdom.Document;
import org.jdom.Element;
import org.jdom.input.SAXBuilder;

/**
 * 创建模板文件
 * @author lhf
 *
 */
public class CreateTemplate {
	public static void main(String[] args) {
		//获取解析xml文件路径
		String path = System.getProperty("user.dir")+"/bin/student.xml";
	    System.out.println(path);
	    File file = new File(path);
	    //创建对象
	    SAXBuilder builder = new SAXBuilder();
	    try {
	    	//解析xml文件
			Document parse = builder.build(file);
			//创建一个Excel
			//创建工作簿
			HSSFWorkbook wb = new HSSFWorkbook();
			//创建sheet
			HSSFSheet sheet = wb.createSheet("Sheet0");
			
			//获取xml文件根节点  <excel></excel>
			Element root = parse.getRootElement();
			//获取模板的名称
			String templateName = root.getAttribute("name").getValue();
			int rownum = 0;  //行号
			int colnum = 0;   //列号
			//设置列宽
			Element colgroup = root.getChild("colgroup");
			setColumnWidth(sheet,colgroup);
			
			/*设置标题 合并单元格
			 *<title>
             *<tr height="16px">
             *<td rowspan="1" colspan="6" value="学生信息导入" />
             *</tr>
             *</title>
			 */
			Element title = root.getChild("title");   //获取title标签<title></title>
			//获取tr标签
			List<Element> trs = title.getChildren("tr");
			//循环tr
			for(int i=0;i<trs.size();i++){
				Element tr = trs.get(i);   //获取tr
				List<Element> tds = tr.getChildren("td");  //获取tr中的td
				//创建一行
				HSSFRow row = sheet.createRow(rownum);
				//为单元格设置样式--居中
				HSSFCellStyle cellStyle = wb.createCellStyle();
				cellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
				//可能有多个td
				for(colnum=0;colnum<tds.size();colnum++){
					Element td = tds.get(colnum);
					//创建单元格
					HSSFCell cell = row.createCell(colnum);
					//获取他的属性
					Attribute rowSpan = td.getAttribute("rowspan");
					Attribute colSpan = td.getAttribute("colspan");
					Attribute value = td.getAttribute("value");
				    //如果value不等于null
					if(value != null){
						String val = value.getValue();
						//将值放入单元格
						cell.setCellValue(val);
						//开始行
						int rspan = rowSpan.getIntValue()-1;
						//合并列
						int cspan = colSpan.getIntValue()-1;
						//设置字体
						HSSFFont font =wb.createFont();
						font.setFontName("仿宋——GB2312");
						//字体加粗
						font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
						//红色字体
						font.setColor(HSSFFont.COLOR_RED);
						//字体高度
						//font.setFontHeight((short)12);
						font.setFontHeightInPoints((short)12);
						//字体加入样式中
						cellStyle.setFont(font);
						//为单元格设置样式
						cell.setCellStyle(cellStyle);
						
						//合并单元格
						sheet.addMergedRegion(new CellRangeAddress(rspan,rspan,0,cspan));
					}
				}
				rownum++;
			}
			
			//设置表头信息
			/* <thead>
               <tr height="16px">
        	   <th value="编号" />
               <th value="姓名" />
               <th value="年龄" />
               <th value="性别" />
               <th value="出生日期" />
               <th value=" 爱好" />            
               </tr>
               </thead>
            */
			//设置表头
			Element thead = root.getChild("thead");
			trs = thead.getChildren("tr");
			for (int i = 0; i < trs.size(); i++) {
				Element tr = trs.get(i);
				//创建一行
				HSSFRow row = sheet.createRow(rownum);
				//获取th的节点信息
				List<Element> ths = tr.getChildren("th");
				for(colnum = 0;colnum < ths.size();colnum++){
					//取出th元素
					Element th = ths.get(colnum);
					//获得th的属性
					Attribute valueAttr = th.getAttribute("value");
					//创建单元格
					HSSFCell cell = row.createCell(colnum);
					if(valueAttr != null){
						String value =valueAttr.getValue();
						//赋值
						cell.setCellValue(value);
					}
				}
				rownum++;
			}
			
			//设置数据区域样式
			Element tbody = root.getChild("tbody");  //获取tbody节点
			Element tr = tbody.getChild("tr");  //获取tr节点
			//获取repeat属性
			int repeat = tr.getAttribute("repeat").getIntValue();
			//获取td
			List<Element> tds = tr.getChildren("td");
			for (int i = 0; i < repeat; i++) {
				//创建行记录
				HSSFRow row = sheet.createRow(rownum);
				//设置单元格
				for(colnum =0 ;colnum < tds.size();colnum++){
					//获取td节点
					Element td = tds.get(colnum);
					//创建单元格
					HSSFCell cell = row.createCell(colnum);
					//设置单元格样式
					setType(wb,cell,td);
				}
				rownum++;
			}
			//把Excel文件保存到本地
			File tempFile = new File("e:/"+templateName + ".xls");
			tempFile.delete();  //存在就删除
			//不存在就创建
			tempFile.createNewFile();
			FileOutputStream stream = FileUtils.openOutputStream(tempFile);
			//写入
			wb.write(stream);
			//关闭
			stream.close();
			
			
			
		} catch (Exception e) {
			e.printStackTrace();
		} 
	}

	/**
	 * 设置单元格样式
	 * @param wb
	 * @param cell
	 * @param td
	 */
	private static void setType(HSSFWorkbook wb, HSSFCell cell, Element td) {
		Attribute typeAttr = td.getAttribute("type");
		String type = typeAttr.getValue();
		HSSFDataFormat format = wb.createDataFormat();
		HSSFCellStyle cellStyle = wb.createCellStyle();
		//判断节点类型
		if("NUMERIC".equalsIgnoreCase(type)){
			cell.setCellType(HSSFCell.CELL_TYPE_NUMERIC);
			Attribute formatAttr = td.getAttribute("format");
			String formatValue = formatAttr.getValue();
			formatValue = StringUtils.isNotBlank(formatValue)? formatValue : "#,##0.00";
			cellStyle.setDataFormat(format.getFormat(formatValue));
		}else if("STRING".equalsIgnoreCase(type)){
			cell.setCellValue("");
			cell.setCellType(HSSFCell.CELL_TYPE_STRING);
			cellStyle.setDataFormat(format.getFormat("@"));
		}else if("DATE".equalsIgnoreCase(type)){
			cell.setCellType(HSSFCell.CELL_TYPE_NUMERIC);
			cellStyle.setDataFormat(format.getFormat("yyyy-m-d"));
		}else if("ENUM".equalsIgnoreCase(type)){
			CellRangeAddressList regions = 
				new CellRangeAddressList(cell.getRowIndex(), cell.getRowIndex(), 
						cell.getColumnIndex(), cell.getColumnIndex());
			Attribute enumAttr = td.getAttribute("format");
			String enumValue = enumAttr.getValue();
			//加载下拉列表内容
			DVConstraint constraint = 
				DVConstraint.createExplicitListConstraint(enumValue.split(","));
			//数据有效性对象
			HSSFDataValidation dataValidation = new HSSFDataValidation(regions, constraint);
			wb.getSheetAt(0).addValidationData(dataValidation);
		}
		cell.setCellStyle(cellStyle);
	}
	/**
	 * 设置列宽
	 * @param sheet
	 * @param colgroup
	 */
	private static void setColumnWidth(HSSFSheet sheet, Element colgroup) {
		//获得每一个孩子
		List<Element> cols = colgroup.getChildren("col");
		for(int i=0;i<cols.size();i++){
			//获取每一个col的设置
			Element col = cols.get(i);
			Attribute width = col.getAttribute("width");
			//得到单位em
			String unit = width.getValue().replaceAll("[0-9,\\.]", "");
			//得到值
			String value = width.getValue().replaceAll(unit, "");
			int v = 0;
			//如果单位为空或者为px
			if(StringUtils.isBlank(unit) || "px".endsWith(unit)){
				//把poi的宽度转化为Excel的宽度
				v = Math.round(Float.parseFloat(value)*37F);
			}else if("em".endsWith(unit)){  //如果单位为em
				v = Math.round(Float.parseFloat(value)*267.5F);
			}
			//设置宽度
			sheet.setColumnWidth(i, v);
			
			
		}
		
	}

}
