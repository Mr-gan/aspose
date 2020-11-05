package aspose;

import com.aspose.cells.*;
import com.aspose.pdf.devices.PngDevice;
import com.aspose.words.Document;
import com.aspose.words.ImageSaveOptions;
import com.aspose.words.SaveFormat;
import org.apache.commons.lang.StringUtils;
import java.io.*;
import java.util.List;
import java.util.*;

/**
 * @Author: gz
 * @Date: 2020/5/27 16:09
 * @Description: aspose2img
 */
public class AsposeUtil {

	/**
	 * 验证License 若不验证会有水印
	 *
	 * @return
	 */
	public boolean getWordLicense() {
		boolean result = false;
		try {
			InputStream is = AsposeUtil.class.getClassLoader()
					.getResourceAsStream("license.xml");
			com.aspose.words.License license = new com.aspose.words.License();
			license.setLicense(is);
			result = true;
		} catch (Exception e) {
			e.printStackTrace();
		}
		return result;
	}

	/**
	 * 验证License 若不验证会有水印
	 *
	 * @return
	 */
	public boolean getExcelLicense() {
		boolean result = false;
		try {
			InputStream is = AsposeUtil.class.getClassLoader()
					.getResourceAsStream("license.xml");
			com.aspose.cells.License license = new com.aspose.cells.License();
			license.setLicense(is);
			result = true;
		} catch (Exception e) {
			e.printStackTrace();
		}
		return result;
	}

	/**
	 * 验证License 若不验证会有水印
	 *
	 * @return
	 */
	public boolean getPdfLicense() {
		boolean result = false;
		try {
			InputStream is = AsposeUtil.class.getClassLoader()
					.getResourceAsStream("license.xml");
			com.aspose.pdf.License license = new com.aspose.pdf.License();
			license.setLicense(is);
			result = true;
		} catch (Exception e) {
			e.printStackTrace();
		}
		return result;
	}

	/**
	 * @param filePath 文件路径
	 * @param outPath  输出路径
	 */
	public void aspose2Img(String filePath, String outPath) throws Exception {
		//注意区分windows和linux /的区分
		String fileName = filePath.substring(filePath.lastIndexOf("\\") + 1, filePath.lastIndexOf("."));
		FileInputStream is = new FileInputStream(filePath);
		if (is == null || StringUtils.isEmpty(fileName) || StringUtils.isEmpty(outPath)) {
			return;
		}
		try {
			if (filePath.endsWith(".doc") || filePath.endsWith(".docx")) {
				getWordLicense();
				//word转图片
				Document doc = new Document(is);
				word2Img(doc, fileName, outPath);
			} else if (filePath.endsWith(".xls") || filePath.endsWith(".xlsx")) {
				getExcelLicense();
				//Excel转图片
				Workbook wb = new Workbook(is);
				excel2Img(wb, fileName, outPath);
			} else if (filePath.endsWith(".pdf")) {
				//pdf转图片
				getPdfLicense();
				com.aspose.pdf.Document document = new com.aspose.pdf.Document(is);
				pdf2Img(document, fileName, outPath);
			}
		} finally {
			if (is != null) {
				try {
					is.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
	}

	/**
	 * Word转图片
	 *
	 * @param doc      文件的Word对象
	 * @param fileName 文件名
	 * @param outPath  目的路径
	 */
	public void word2Img(Document doc, String fileName, String outPath) {
		try {
			//设置图片输出格式
			ImageSaveOptions iso = new ImageSaveOptions(SaveFormat.PNG);
			//图片数量为Word文档页数
			int pageCount = doc.getPageCount();
			for (int i = 1; i <= pageCount; i++) {
				//图片输出位置
				File file = new File(outPath + fileName + "-" + i + ".png");
				FileOutputStream os = new FileOutputStream(file);
				iso.setPageIndex(i - 1);
				//文档转图片
				doc.save(os, iso);
				os.close();
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	/**
	 * excel转图片
	 *
	 * @param wb
	 * @param outPath
	 */
	public void excel2Img(Workbook wb, String fileName, String outPath) {
		try {
			//excel sheet数量
			int sheetCount = wb.getWorksheets().add();
			for (int i = 1; i <= sheetCount; i++) {
				Worksheet sheet = wb.getWorksheets().get(i - 1);
				if (sheet.getCells().getMaxDataColumn() != -1) {
					//有数据的sheet才输出
					//设置图片输出格式
					ImageOrPrintOptions iop = new ImageOrPrintOptions();
					iop.setChartImageType(ImageFormat.getPng());
					iop.setCellAutoFit(true);
					iop.setOnePagePerSheet(true);
					SheetRender render = new SheetRender(sheet, iop);
					String outName = outPath + fileName + "-" + i + ".png";
					render.toImage(0, outName);
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
	}


	/**
	 * pdf转图片
	 *
	 * @param document
	 * @param outPath
	 */
	public void pdf2Img(com.aspose.pdf.Document document, String fileName, String outPath) {
		try {
			//图片数量为PDF文档页数
			int pageCount = document.getPages().size();
			PngDevice pngDevice = new PngDevice();
			for (int i = 1; i <= pageCount; i++) {
				//图片输出位置
				File file = new File(outPath + fileName + "-" + i + ".png");
				FileOutputStream os = new FileOutputStream(file);
				//文档转图片
				pngDevice.process(document.getPages().get_Item(i), os);
				os.close();
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public static void main(String[] args) throws Exception {
		AsposeUtil asposeUtil = new AsposeUtil();
		String outPath = "C:\\Users\\Administrator\\Desktop\\aspose\\";
//		String filePath = "C:\\Users\\Administrator\\Desktop\\222.pdf";
		String filePath = "C:\\Users\\Administrator\\Desktop\\111.doc";
		asposeUtil.aspose2Img(filePath,outPath);

	}
}
