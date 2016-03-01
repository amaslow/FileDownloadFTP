package filedownloadftp;

import java.io.BufferedOutputStream;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.OutputStream;
import java.net.SocketTimeoutException;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.Iterator;
import java.util.List;
import org.apache.commons.net.ftp.FTP;
import org.apache.commons.net.ftp.FTPClient;
import org.apache.commons.net.ftp.FTPFile;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.*;

public class FileDownloadFTP {

    static String excelSource = "G:\\CM\\Category Management Only\\_S0000_Trade marketing\\Pictures Spaceman\\SAP_EAN.xlsx";
    static String excelOutput = "G:\\Product Content\\Sales & Content tools\\New product picture log file\\New product picture log file.xlsx";
    static String folderDest = "G:\\Product Content\\PRODUCTS\\";
    static ArrayList<String> sap = new ArrayList<String>();

    public static void main(String[] args) throws IOException, ParseException, NullPointerException, SocketTimeoutException {
        String src = "/";
        File dst = new File(folderDest);
        FileWriter fw = new FileWriter("H:/Logs/FileDownloadFtp.log", true);
        BufferedWriter bw = new BufferedWriter(fw);

        String server = "ftp.tristar.eu";
        int port = 21;
        String user = "TransferSLfotoDB";
        String pass = "5bnV2iss";

        FTPClient ftpClient = new FTPClient();
        try {
            ftpClient.connect(server, port);
            ftpClient.login(user, pass);
            ftpClient.enterLocalPassiveMode();
            ftpClient.setFileType(FTP.BINARY_FILE_TYPE);
            System.out.println("Connected to FTP...");
            bw.newLine();
            bw.newLine();
            bw.write("Connected to FTP...");
            bw.newLine();

            listDirectory(bw, ftpClient, src, "", dst);

        } catch (IOException ex) {
            System.out.println("Error: " + ex.getMessage());
            bw.newLine();
            bw.write("Error: " + ex.getMessage());
            ex.printStackTrace();
        } finally {
            try {
                if (ftpClient.isConnected()) {
                    ftpClient.logout();
                    ftpClient.disconnect();
                    System.out.println("Disconnected");
                    bw.newLine();
                    bw.write("Disconnected");
                    System.out.println("----------------------------------------------");
                    bw.newLine();
                    bw.write("----------------------------------------------");
                    for (int i = 0; i < sap.size(); i++) {
                        createLog(sap.get(i));
                    }

                }
            } catch (IOException ex) {
                ex.printStackTrace();
            }
        }
    }

    static void listDirectory(BufferedWriter bw, FTPClient ftpClient, String parentDir, String currentDir, File destDir) throws IOException, NullPointerException {
        String dirToList = parentDir;
        if (!currentDir.equals("")) {
            dirToList += "/" + currentDir;
        }

        FTPFile[] subFiles = ftpClient.listFiles(dirToList);
        if (subFiles != null && subFiles.length > 0) {
            for (FTPFile aFile : subFiles) {
                String currentFileName = aFile.getName();
                if (currentFileName.equals(".") || currentFileName.equals("..") || currentFileName.equals("Thumbs.db")) {
                    continue;
                }
                if (aFile.isDirectory()) {
                    if (!currentFileName.equals("output")) {
                        String material = currentFileName.substring(0, 2) + "." + currentFileName.substring(2, 5) + "." + currentFileName.substring(5, 7);
                        sap.add(material);
                    }
                    listDirectory(bw, ftpClient, dirToList, currentFileName, destDir);

                } else {
                    if (!currentFileName.toLowerCase().endsWith("jpg")) {
                        continue;
                    }
                    copyFiles(bw, ftpClient, currentFileName, destDir);
                }
            }
//            for (int i = 0; i < sap.size(); i++) {
//                createLog(sap.get(i));
//            }
            for (FTPFile dFile : subFiles) {
                String currentFileName = dFile.getName();
                if (currentFileName.equals(".") || currentFileName.equals("..")) {
                    continue;
                }
                if (!dFile.isDirectory()) {
                    boolean successDelete = ftpClient.deleteFile(dirToList + "/" + dFile.getName());
                    if (successDelete) {
                        System.out.println("File " + dFile.getName() + " has been deleted.");
                        bw.newLine();
                        bw.write("File " + dFile.getName() + " has been deleted.");
                    }
                } else {
                    boolean successDelete = ftpClient.removeDirectory(dirToList + "/" + dFile.getName());
                    if (successDelete) {
                        System.out.println("Folder " + currentFileName + " has been deleted.");
                        bw.newLine();
                        bw.write("Folder " + currentFileName + " has been deleted.");
                    }
                }

            }

        }
    }

    static void copyFiles(BufferedWriter bw, FTPClient ftpClient, String currentFileName, File destDir) throws FileNotFoundException, IOException {
        String currentDirName = currentFileName.substring(3, 10);
        String remoteFile = "/" + currentDirName + "/output/" + currentFileName;
        File destination = new File(destDir + "/" + currentDirName + "/");
        if (!destination.exists()) {
            destination.mkdirs();
        }
        OutputStream outputStream = new BufferedOutputStream(new FileOutputStream(destination + "/" + currentFileName));
        boolean successCopy = ftpClient.retrieveFile(remoteFile, outputStream);
        outputStream.close();
        if (successCopy) {
            System.out.println("File " + currentFileName + " has been downloaded successfully.");
            bw.newLine();
            bw.write("File " + currentFileName + " has been downloaded successfully.");
        } else {
            System.out.println("File " + currentFileName + " not downloaded.");
            bw.newLine();
            bw.write("File " + currentFileName + " not downloaded.");
        }
    }

    static void createLog(String material) throws IOException, FileNotFoundException, NullPointerException {
        int week = Calendar.getInstance().get(Calendar.WEEK_OF_YEAR);
        DateFormat dateFormater = new SimpleDateFormat("dd-MM-yyyy");
        String modDate = dateFormater.format(new Date());
        String pictureName = null;
        String pictureName2 = folderDest + material.replace(".", "") + "\\LR_" + material.replace(".", "") + "_2.jpg";
        String pictureName3 = folderDest + material.replace(".", "") + "\\LR_" + material.replace(".", "") + "_3.jpg";
        File pic2 = new File(pictureName2);
        File pic3 = new File(pictureName3);
        if (pic2.exists()) {
            pictureName = pictureName2;
        } else if (pic3.exists()) {
            pictureName = pictureName3;
        } else {
            pictureName = null;
        }
        FileInputStream fis = null;
        fis = new FileInputStream(new File(excelOutput));

        XSSFWorkbook wb = new XSSFWorkbook(fis);
        XSSFSheet sheet = wb.getSheetAt(0);
        CellStyle style = wb.createCellStyle();
        style.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);

        int last = (sheet.getLastRowNum() + 1);

        XSSFRow row = sheet.createRow(last);
        row.setHeight((short) 1000);
        XSSFCell weekCell = row.createCell(0);
        weekCell.setCellValue(week);
        weekCell.setCellStyle(style);
        XSSFCell modDateCell = row.createCell(1);
        modDateCell.setCellValue(modDate);
        modDateCell.setCellStyle(style);
        XSSFCell materialCell = row.createCell(2);
        materialCell.setCellValue(material);
        materialCell.setCellStyle(style);
        XSSFCell descrCell = row.createCell(4);
        descrCell.setCellValue(getDescrBySap(material));
        descrCell.setCellStyle(style);
        if (pictureName != null) {
            /* Read input PNG / JPG Image into FileInputStream Object*/
            FileInputStream image = new FileInputStream(pictureName);
            /* Convert picture to be added into a byte array */
            byte[] bytes = IOUtils.toByteArray(image);
            /* Add Picture to Workbook, Specify picture type as PNG and Get an Index */
            int my_picture_id = wb.addPicture(bytes, XSSFWorkbook.PICTURE_TYPE_JPEG);
            /* Close the InputStream. We are ready to attach the image to workbook now */
            image.close();
            /* Create the drawing container */
            XSSFDrawing drawing = sheet.createDrawingPatriarch();
            /* Create an anchor point */
            XSSFClientAnchor my_anchor = new XSSFClientAnchor();
            /* Define top left corner, and we can resize picture suitable from there */
            int col1 = 3, row1 = last;
            my_anchor.setCol1(col1);
            my_anchor.setRow1(row1);
            my_anchor.setDx1(0);
            my_anchor.setDy1(0);
            my_anchor.setCol2((short) ++col1);
            my_anchor.setRow2(++row1);
            my_anchor.setDx2(0);
            my_anchor.setDy2(0);
            drawing.createPicture(my_anchor, my_picture_id);
        }
        //System.out.println("rownr: " + last + " week: " + week + " modDate: " + modDate + " material: " + material + " description: " + getDescrBySap(material));
        fis.close();
        FileOutputStream fos = new FileOutputStream(new File(excelOutput));
        wb.write(fos);
        fos.close();
    }

    public static String getDescrBySap(String material) throws NullPointerException, IOException {
        List sheetData = new ArrayList();
        FileInputStream fis_excel = null;
        try {
            fis_excel = new FileInputStream(excelSource);
            XSSFWorkbook wb_excel = new XSSFWorkbook(fis_excel);
            XSSFSheet sheet_excel = wb_excel.getSheetAt(0);
            Iterator rows = sheet_excel.rowIterator();
            while (rows.hasNext()) {
                XSSFRow row = (XSSFRow) rows.next();
                List data = new ArrayList();
                data.add(row.getCell(0));
                if (row.getCell(2) != null) {
                    data.add(row.getCell(2));
                } else {
                    data.add("");
                }
                sheetData.add(data);
            }
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (fis_excel != null) {
                fis_excel.close();
            }
        }
        for (int i = 0; i < sheetData.size(); i++) {
            List list = (List) sheetData.get(i);
            if (material.equals(list.get(0).toString())) {
                String descr = list.get(1).toString();
                return descr;
            }
        }
        return null;
    }

}
