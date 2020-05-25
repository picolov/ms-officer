package com.picolov;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import net.sf.mpxj.MPXJException;
import net.sf.mpxj.ProjectFile;
import net.sf.mpxj.Task;
import net.sf.mpxj.reader.ProjectReader;
import net.sf.mpxj.reader.UniversalProjectReader;

public class App {
    public static void main(final String[] args) {
        String inputFolder = ".";
        String templateFolder = ".";
        String outputFolder = ".";
        if (args.length > 0)
            inputFolder = args[0];
        if (args.length > 1)
            templateFolder = args[1];
        if (args.length > 2)
            outputFolder = args[2];
        List<Map<String, Object>> rowLoopPlaceholders;
        try {
            final File inFolder = new File(inputFolder);
            final File temfolder = new File(templateFolder);

            final List<String> mppResult = new ArrayList<>();
            search(".*\\.mpp", inFolder, mppResult);
            final List<String> templateResult = new ArrayList<>();
            search(".*_template\\.xlsx", temfolder, templateResult);

            for (final String inputMppFile : mppResult) {
                rowLoopPlaceholders = new ArrayList<>();
                System.out.println("processing mpp : " + inputMppFile);
                String filename = inputMppFile.split("\\.")[0];
                String fileResult = outputFolder + "/" + filename + ".xlsx";
                String fileTemplate = templateFolder + "/" + "template.xlsx";
                String fileTemplateBak = templateFolder + "/" + "BAK_template.xlsx";
                for (int loop = 0; loop < templateResult.size(); loop++) {
                    String templateName = templateResult.get(loop);
                    String templateNamePattern = templateName.substring(0, templateName.length() - 14);
                    if (filename.startsWith(templateNamePattern)) {
                        fileTemplate = templateFolder + "/" + templateName;
                        fileTemplateBak = templateFolder + "/" + "BAK_" + templateName;
                        break;
                    }
                }
                System.out.println("using template : " + fileTemplate);
                // read mpp
                ProjectReader reader = new UniversalProjectReader();
                ProjectFile project = reader.read(inputFolder + "/" + inputMppFile);
                List<Map<String, Object>> taskList = new ArrayList<>();
                listHierarchy(project, taskList);
                // read template placeholder
                Workbook workbookR = WorkbookFactory.create(new File(fileTemplate));
                Sheet sheetR = workbookR.getSheetAt(0);
                Iterator<Row> iterator = sheetR.iterator();
                while (iterator.hasNext()) {
                    Row currentRow = iterator.next();
                    Iterator<Cell> cellIterator = currentRow.iterator();
                    while (cellIterator.hasNext()) {
                        Cell currentCell = cellIterator.next();
                        // getCellTypeEnum shown as deprecated for version 3.15
                        // getCellTypeEnum ill be renamed to getCellType starting from version 4.0
                        if (currentCell.getCellType() == CellType.STRING
                                && currentCell.getStringCellValue().startsWith("{{#")) {
                            String value = currentCell.getStringCellValue();
                            Map<String, Object> placeholder = new HashMap<>();
                            String key = value.substring(3, value.length() - 2);
                            if (key.contains("|")) {
                                String[] keyToken = key.split("\\|");
                                key = keyToken[0];
                                for (int i = 1; i < keyToken.length; i++) {
                                    String addOn = keyToken[i];
                                    if (addOn.equals("AUTOSIZE")) {
                                        placeholder.put("autosize", true);
                                    }
                                }
                            }
                            placeholder.put("key", key);
                            placeholder.put("row", currentCell.getRowIndex());
                            placeholder.put("col", currentCell.getColumnIndex());
                            rowLoopPlaceholders.add(placeholder);
                        }
                    }
                }
                System.out.println("loopPlaceholder: " + rowLoopPlaceholders);
                workbookR.close();
                // preserve template
                Files.deleteIfExists(Paths.get(fileTemplateBak));
                Files.copy(Paths.get(fileTemplate), Paths.get(fileTemplateBak));
                // write to xlsx
                Workbook workbookW = WorkbookFactory.create(new File(fileTemplate));
                Sheet sheetW = workbookW.getSheetAt(0);
                for (Map<String, Object> placeholder : rowLoopPlaceholders) {
                    int rowIdx = (Integer) placeholder.get("row");
                    int colIdx = (Integer) placeholder.get("col");
                    String key = (String) placeholder.get("key");
                    // ┗┣━━
                    for (Map<String, Object> task : taskList) {
                        Row row = sheetW.getRow(rowIdx);
                        if (row == null)
                            row = sheetW.createRow(rowIdx);
                        Cell cell = row.getCell(colIdx);
                        if (cell == null)
                            cell = row.createCell(colIdx);
                        switch (key) {
                            case "wbs":
                                cell.setCellValue((String) task.get(key));
                                break;
                            case "taskName":
                                String value = (String) task.get(key);
                                int indent = (Integer) task.get("indent");
                                for (int i = 0; i < indent; i++) {
                                    value = "   " + value;
                                }
                                cell.setCellValue(value);
                                break;
                            case "startDate":
                                cell.setCellValue((Date) task.get(key));
                                break;
                            case "finishDate":
                                cell.setCellValue((Date) task.get(key));
                                break;
                            default:
                                cell.setCellValue("KEY -" + key + " NOT FOUND");
                        }
                        rowIdx++;
                    }

                }

                for (Map<String, Object> placeholder : rowLoopPlaceholders) {
                    if (placeholder.containsKey("autosize") && (Boolean) placeholder.get("autosize")) {
                        int colIdx = (Integer) placeholder.get("col");
                        sheetW.autoSizeColumn(colIdx);
                    }
                }
                FileOutputStream fileOut = new FileOutputStream(fileResult, false);
                workbookW.write(fileOut);
                fileOut.close();
                workbookW.close();

                // revert back the template, because when closing somehow POI is writing the
                // template also
                Files.deleteIfExists(Paths.get(fileTemplate));
                new File(fileTemplateBak).renameTo(new File(fileTemplate));
            }
        } catch (final MPXJException e) {
            e.printStackTrace();
        } catch (final EncryptedDocumentException e) {
            e.printStackTrace();
        } catch (final IOException e) {
            e.printStackTrace();
        }
    }

    public static void listHierarchy(final ProjectFile file, List<Map<String, Object>> taskList) {
        for (final Task task : file.getChildTasks()) {
            Map<String, Object> taskMap = new HashMap<>();
            taskMap.put("uid", task.getUniqueID());
            taskMap.put("taskName", task.getName());
            taskMap.put("startDate", task.getStart());
            taskMap.put("finishDate", task.getFinish());
            taskMap.put("wbs", task.getWBS());
            taskMap.put("indent", 0);
            taskMap.put("isLeaf", false);
            taskList.add(taskMap);
            listHierarchy(task, 0, taskList);
        }
    }

    private static void listHierarchy(final Task task, final int parentDeep, List<Map<String, Object>> taskList) {
        for (final Task child : task.getChildTasks()) {
            Map<String, Object> taskMap = new HashMap<>();
            taskMap.put("uid", child.getUniqueID());
            taskMap.put("taskName", child.getName());
            taskMap.put("startDate", child.getStart());
            taskMap.put("finishDate", child.getFinish());
            taskMap.put("wbs", child.getWBS());
            taskMap.put("indent", parentDeep + 1);
            if (child.getChildTasks().size() == 0)
                taskMap.put("isLeaf", true);
            else
                taskMap.put("isLeaf", false);
            taskList.add(taskMap);
            listHierarchy(child, parentDeep + 1, taskList);
        }
    }

    public static void search(final String pattern, final File folder, final List<String> result) {
        for (final File f : folder.listFiles()) {
            // if (f.isDirectory()) {
            // search(pattern, f, result);
            // }
            if (f.isFile()) {
                if (f.getName().matches(pattern)) {
                    // result.add(f.getAbsolutePath());
                    result.add(f.getName());
                }
            }
        }
    }
}
