/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package ie.enovation;

import org.apache.commons.cli.CommandLine;
import org.apache.commons.cli.CommandLineParser;
import org.apache.commons.cli.HelpFormatter;
import org.apache.commons.cli.Options;
import org.apache.commons.cli.PosixParser;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileReader;
import java.io.FileWriter;
import java.util.ArrayList;
import java.util.List;
import java.util.Properties;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

/**
 * Converts an Excel file into a set of items to import into DSpace
 * @author pvillega
 */
public class Main {

    public static Properties props;

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        try {
            //load properties
            props = new Properties();
            props.load(new FileReader(new File("columns.properties")));

            CommandLineParser parser = new PosixParser();
            Options options = new Options();
            options.addOption("i", "input", true, "Excel file used as input");
            options.addOption("o", "output", true, "Folder where the structure will be placed");

            CommandLine line = parser.parse(options, args);

            if (line.hasOption('h')) {
                HelpFormatter myhelp = new HelpFormatter();
                myhelp.printHelp("Excel2Dspace\n", options);
                System.exit(0);
            }

            if (!line.hasOption('i') || !line.hasOption('o')) {
                System.out.println("Mandatory parameters are missing!");
                HelpFormatter myhelp = new HelpFormatter();
                myhelp.printHelp("Excel2Dspace\n", options);
                System.exit(-1);
            }

            File input = new File(line.getOptionValue('i'));
            if (!input.exists() || !input.isFile()) {
                System.out.println("The input parameter must be an existing Excel file");
                System.exit(-1);
            }

            File output = new File(line.getOptionValue('o'));
            if (output.exists() && !output.isDirectory()) {
                System.out.println("The output parameter must be a folder");
                System.exit(-1);
            }
            if (!output.exists()) {
                if (!output.mkdirs()) {
                    System.out.println("Couldn't create the output folder");
                    System.exit(-1);
                }
            }

            //#1 read the excel file
            System.out.println("#1 - Opening " + input.getAbsolutePath());
            HSSFWorkbook excel = new HSSFWorkbook(new FileInputStream(input));

            List<Item> items = new ArrayList<Item>();

            //reads all the sheets in the excel file and stores the rows as items
            for (int k = 0; k < excel.getNumberOfSheets(); k++) {
                HSSFSheet sheet = excel.getSheetAt(k);
                int rows = sheet.getPhysicalNumberOfRows();
                System.out.println(" >> Sheet " + k + " \"" + excel.getSheetName(k) + "\" has " + rows + " row(s).");
                for (int r = 0; r < rows; r++) {
                    Item item = new Item();
                    HSSFRow row = sheet.getRow(r);
                    if (row == null) {
                        continue;
                    }

                    int cells = row.getPhysicalNumberOfCells();
                    System.out.println("\nROW " + row.getRowNum() + " has " + cells + " cell(s).");

                    //TODO: this can be improved to become a more generic converter reading from a mappings file
                    // column number - dc field
                    for(int c = 0; c < cells; c++) {
                        System.out.println("\nCell "+c);
                        HSSFCell cell = row.getCell(c);
                        if(cell != null) { //filter errors in excel file
                            item.add(props.getProperty("column.metadata."+c), getCellValue(cell),props.getProperty("column.type."+c));
                        }
                    }

                    items.add(item);
                }
            }


            //#2 create output folder structure (or empty if existing)
            File parent = new File(line.getOptionValue('o') + File.separator + "archive_folder");
            System.out.println("#2 - Creating/Emptying " + parent.getAbsolutePath());
            if (!parent.exists()) {
                if (!parent.mkdirs()) {
                    System.out.println("Couldn't create the archives folder");
                    System.exit(-1);
                }
            } else {
                //if it exists, we clean the contents
                emptyFolder(parent);
            }

            //#3 iterate and create items
            System.out.println("#3 - Creating the items");
            int i = 0;
            for (Item it : items) {
                File folder = new File(parent.getAbsolutePath() + File.separator + "item_" + i);
                folder.mkdir();

                File xml = new File(folder.getAbsolutePath() + File.separator + "dublin_core.xml");
                xml.createNewFile();
                FileWriter fw = new FileWriter(xml);
                fw.write(it.toXML());
                fw.close();

                File content = new File(folder.getAbsolutePath() + File.separator + "contents");
                content.createNewFile();

                i++;
            }

        } catch (Exception e) {
            e.printStackTrace();
        }

    }

    private static String getCellValue(HSSFCell cell) {
        String value = "";
        switch (cell.getCellType()) {

            case HSSFCell.CELL_TYPE_FORMULA:
                value = cell.getCellFormula().toString();
                break;

            case HSSFCell.CELL_TYPE_NUMERIC:
                value = Double.toString(cell.getNumericCellValue());
                break;

            case HSSFCell.CELL_TYPE_STRING:
                value = cell.getStringCellValue();
                break;

            default:
        }
        return value;
    }

    private static void emptyFolder(File folder) {
        File[] oldItems = folder.listFiles();
        for (File f : oldItems) {
            if (f.isDirectory()) {
                emptyFolder(f);
            }
            f.delete();
        }
    }
}
