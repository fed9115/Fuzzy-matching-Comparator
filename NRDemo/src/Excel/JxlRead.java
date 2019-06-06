package Excel;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;

import javafx.util.Pair;
import jxl.*;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

public class JxlRead {

	public static void excelComp(File f1, File f2, String savePath) {
		// Setting 2 array lists to store data which need to be compared
		ArrayList<String> al1 = new ArrayList<>();
		ArrayList<String> al2 = new ArrayList<>();
		ArrayList<Pair<String, String>> result = new ArrayList<>();

		// Reading two files using jxl
		try {
			// Creating two books which represent two files
			Workbook book1 = Workbook.getWorkbook(f1);
			Workbook book2 = Workbook.getWorkbook(f2);

			// Extracting the sheets from the 2 books
			Sheet s1 = book1.getSheet(0);
			Sheet s2 = book2.getSheet(0);

			// Getting the data which are strings from column B's from 2 sheets, these data
			// are which we need to do comparisons, storing them into two array lists
			for (int i = 1; i < s1.getRows(); ++i) {
				// Since we need to get column B which index is 1 and row is i, Cell c is the
				// value of the current data we iterate from the current sheet
				Cell c = s1.getCell(1, i);
				al1.add(c.getContents());
			}
			for (int i = 1; i < s2.getRows(); ++i) {
				// Since we need to get column B which index is 1 and row is i, Cell c is the
				// value of the current data we iterate from the current sheet
				Cell c = s2.getCell(1, i);
				al2.add(c.getContents());
			}

			/**
			 * ALGORITHM 1 - CHECK IF THE CODES OF THE LINE IN TWO ARRAYLISTS ARE THE SAME
			 */

			// Creating an array list of maps to store the lines' codes which map to its
			// name if exist
			ArrayList<Pair<Integer, String>> lineCode1 = new ArrayList<>();
			ArrayList<Pair<Integer, String>> lineCode2 = new ArrayList<>();

			for (int i = 0; i < al1.size(); ++i) {
				// Splitting the string (name of lines into several parts: ints and strings)
				String[] parts = al1.get(i).split("(?<=\\D)(?=\\d)");
				for (String part : parts) {
					if (isNumber(part)) {
						// Check if the part which is a number but is not 110 (since 110 is always the
						// number of voltage but not code)
						if (Integer.parseInt(part) != 110) {
							Pair<Integer, String> p = new Pair<>(Integer.parseInt(part), part);
							// Adding the pair to the new array list, removing the original string from the
							// array list al1
							lineCode1.add(p);
							al1.remove(i);
							break;
						}
					}
				}
			}

			for (int i = 0; i < al2.size(); ++i) {
				// Splitting the string (name of lines into several parts: ints and strings)
				String[] parts = al2.get(i).split("(?<=\\D)(?=\\d)");
				for (String part : parts) {
					if (isNumber(part)) {
						// Check if the part which is a number but is not 110 (since 110 is always the
						// number of voltage but not code)
						if (Integer.parseInt(part) != 110) {
							Pair<Integer, String> p = new Pair<>(Integer.parseInt(part), part);
							// Adding the pair to the new array list, removing the original string from the
							// array list al1
							lineCode2.add(p);
							al2.remove(i);
							break;
						}
					}
				}
			}

			// If there is a code of line in lineCode2 matching to one in lineCode1, put the
			// pair of strings (which is the name of the line) into a newly generated array
			// list
			for (int i = 0; i < lineCode1.size(); ++i) {
				for (int j = 0; j < lineCode2.size(); ++j) {
					if (lineCode1.get(i).getKey() == lineCode2.get(j).getKey()) {
						Pair<String, String> p = new Pair<>(lineCode1.get(i).getValue(), lineCode2.get(j).getValue());
						result.add(p);
					}
				}
			}

			/**
			 * ALGORITHM 2 - CHECK REMAINING STRINGS IN THE 2 ARRAY LISTS, MATCHING TWO
			 * STRINGS WITH HIGHEST REPITITION RATE
			 */
			for (int i = 0; i < al1.size(); ++i) {
				ArrayList<Pair<String, Integer>> repetitionRate = new ArrayList<>();
				for (int j = 0; j < al2.size(); ++j) {
					int count = 0;
					// check the character one by one, if there's one in the al2 matching the
					// current char in al1, increment count
					for (char c1 : al1.get(i).toCharArray()) {
						for (char c2 : al2.get(j).toCharArray()) {
							// if (c1 == 'Ⅰ' || c1 == 'Ⅱ' || c1 == 'II' || c2 == 'Ⅰ' || c2 == 'Ⅱ' || c2 ==
							// 'II') Ⅳ Ⅰ Ⅱ Ⅲ Ⅰ
							// Convert Roman numerical to Chinese characters
							if (c1 == 'Ⅰ')
								c1 = '一';
							if (c1 == 'Ⅱ')
								c1 = '二';
							if (c1 == 'Ⅲ')
								c1 = '三';
							if (c1 == 'Ⅳ')
								c1 = '四';
							if (c2 == 'Ⅰ')
								c2 = '一';
							if (c2 == 'Ⅱ')
								c2 = '二';
							if (c2 == 'Ⅲ')
								c2 = '三';
							if (c2 == 'Ⅳ')
								c2 = '四';

							if (c1 == c2)
								count++;
						}
					}
					repetitionRate.add(new Pair<>(al2.get(j), count));
				}
				// Finding the string with largest count which represents the highest repetition
				// rate
				Pair<String, Integer> max = repetitionRate.get(0);
				for (int k = 0; k < repetitionRate.size(); ++k) {
					if (repetitionRate.get(k).getValue() > max.getValue()) {
						max = repetitionRate.get(k);
					}
				}
				String maxKey = max.getKey();
				result.add(new Pair<String, String>(al1.get(i), maxKey));
			}

			// Export the array list of Pairs to an Excel table
			WritableWorkbook book = null;
			File fileToSave = new File(savePath);

			fileToSave.createNewFile();
			// Create an instance of excle book
			book = Workbook.createWorkbook(fileToSave);

			// Create a sheet in the book
			WritableSheet sheet = book.createSheet("sheet1", 0);
			for (int i = 0; i < result.size(); ++i) {
				Label l1 = new Label(0, i, result.get(i).getKey());
				Label l2 = new Label(1, i, result.get(i).getValue());

				sheet.addCell(l1);
				sheet.addCell(l2);

				book.write();
			}

		} catch (BiffException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (RowsExceededException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (WriteException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	/**
	 * Helper method for me to check if the string is a number
	 * 
	 * @param part
	 * @return
	 */
	private static boolean isNumber(String part) {
		try {
			Integer.parseInt(part);
			return true;
		} catch (NumberFormatException e) {
			return false;
		}
	}

	public static void main(String[] args) {
		File f1 = new File(args[0]);
		File f2 = new File(args[1]);
		excelComp(f1, f2, args[2]);
	}

}
