import java.awt.Color;
import java.awt.Image;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Scanner;
import java.io.IOException;

import org.apache.poi.ss.usermodel.ConditionalFormattingRule;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.SheetConditionalFormatting;
import org.apache.poi.ss.usermodel.ExtendedColor;
import org.apache.poi.ss.usermodel.ColorScaleFormatting;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.ConditionalFormattingThreshold.RangeType;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.imageio.ImageIO;

public class MainApp {

	public static void main(String[] args) {

		Scanner in = new Scanner(System.in);
		System.out.println("Please enter your image's full path:");
		Path imagePath = Paths.get(in.nextLine().replace("\"", ""));
		in.close();
		Workbook wb = new XSSFWorkbook();
		

		try {
			BufferedImage bi = ImageIO.read(imagePath.toFile());
			
			if (bi.getWidth() > 223) {
				Image resizedImage = bi.getScaledInstance(223, (int) (223.0/bi.getWidth() * bi.getHeight()), Image.SCALE_DEFAULT);
				bi = new BufferedImage(resizedImage.getWidth(null), resizedImage.getHeight(null), BufferedImage.TYPE_INT_ARGB);
				bi.createGraphics().drawImage(resizedImage, 0, 0, null);
				bi.createGraphics().dispose();
			}
			
			
			File excelFile = new File(imagePath.getParent().toString(), imagePath.getFileName().toString().replaceFirst("\\..*", "") + ".xlsx");
			excelFile.createNewFile();
			
			OutputStream excelFileStream = new FileOutputStream(excelFile);
			
			
			
//			OutputStream excelFile = new FileOutputStream("image.xlsx");
			
			Sheet sheet = wb.createSheet("Sheet1");

			String red = "FFFF0000";
			String green = "FF00FF00";
			String blue = "FF0000FF";
			String black = "FF000000";

			SheetConditionalFormatting conditionalFormatting = sheet.getSheetConditionalFormatting();

			ConditionalFormattingRule redRule = conditionalFormatting.createConditionalFormattingColorScaleRule();

			ColorScaleFormatting colorScaleFormattingRed = redRule.getColorScaleFormatting();

			colorScaleFormattingRed.setNumControlPoints(2);
			colorScaleFormattingRed.getThresholds()[0].setRangeType(RangeType.NUMBER);
			colorScaleFormattingRed.getThresholds()[0].setValue(0d);
			colorScaleFormattingRed.getThresholds()[1].setRangeType(RangeType.NUMBER);
			colorScaleFormattingRed.getThresholds()[1].setValue(255d);
			((ExtendedColor)colorScaleFormattingRed.getColors()[0]).setARGBHex(black);
			((ExtendedColor)colorScaleFormattingRed.getColors()[1]).setARGBHex(red);

			ConditionalFormattingRule greenRule = conditionalFormatting.createConditionalFormattingColorScaleRule();

			ColorScaleFormatting colorScaleFormattingGreen = greenRule.getColorScaleFormatting();

			colorScaleFormattingGreen.setNumControlPoints(2);
			colorScaleFormattingGreen.getThresholds()[0].setRangeType(RangeType.NUMBER);
			colorScaleFormattingGreen.getThresholds()[0].setValue(0d);
			colorScaleFormattingGreen.getThresholds()[1].setRangeType(RangeType.NUMBER);
			colorScaleFormattingGreen.getThresholds()[1].setValue(255d);
			((ExtendedColor)colorScaleFormattingGreen.getColors()[0]).setARGBHex(black);
			((ExtendedColor)colorScaleFormattingGreen.getColors()[1]).setARGBHex(green);

			ConditionalFormattingRule blueRule = conditionalFormatting.createConditionalFormattingColorScaleRule();

			ColorScaleFormatting colorScaleFormattingBlue = blueRule.getColorScaleFormatting();

			colorScaleFormattingBlue.setNumControlPoints(2);
			colorScaleFormattingBlue.getThresholds()[0].setRangeType(RangeType.NUMBER);
			colorScaleFormattingBlue.getThresholds()[0].setValue(0d);
			colorScaleFormattingBlue.getThresholds()[1].setRangeType(RangeType.NUMBER);
			colorScaleFormattingBlue.getThresholds()[1].setValue(255d);
			((ExtendedColor)colorScaleFormattingBlue.getColors()[0]).setARGBHex(black);
			((ExtendedColor)colorScaleFormattingBlue.getColors()[1]).setARGBHex(blue);

			CellRangeAddress[] redRegions = new CellRangeAddress[bi.getHeight()];
			CellRangeAddress[] blueRegions = new CellRangeAddress[bi.getHeight()];
			CellRangeAddress[] greenRegions = new CellRangeAddress[bi.getHeight()];

			int width = bi.getWidth();

			for (int y = 0; y < bi.getHeight(); y++) {
				int r = y * 3;
				Row row1 = sheet.createRow(r);
				Row row2 = sheet.createRow(r + 1);
				Row row3 = sheet.createRow(r + 2);

				for (int x = 0; x < width; x++) {

					Color color = new Color(bi.getRGB(x, y));

					row1.createCell(x).setCellValue(color.getRed());
					row2.createCell(x).setCellValue(color.getGreen());
					row3.createCell(x).setCellValue(color.getBlue());
				}

				redRegions[y] = new CellRangeAddress(r, r, 0, width);
				greenRegions[y] = new CellRangeAddress(r + 1, r + 1, 0, width);
				blueRegions[y] = new CellRangeAddress(r + 2, r + 2, 0, width);
				
			}
			bi.flush();

			conditionalFormatting.addConditionalFormatting(redRegions, redRule);
			conditionalFormatting.addConditionalFormatting(greenRegions, greenRule);
			conditionalFormatting.addConditionalFormatting(blueRegions, blueRule);

			wb.write(excelFileStream);
			wb.close();
			
		} catch (IOException e) {
			e.printStackTrace();	
		}
		
	}
	
}
