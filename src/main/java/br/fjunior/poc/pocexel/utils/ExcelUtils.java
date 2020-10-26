package br.fjunior.poc.pocexel.utils;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xddf.usermodel.chart.AxisCrosses;
import org.apache.poi.xddf.usermodel.chart.AxisPosition;
import org.apache.poi.xddf.usermodel.chart.ChartTypes;
import org.apache.poi.xddf.usermodel.chart.LegendPosition;
import org.apache.poi.xddf.usermodel.chart.XDDFCategoryAxis;
import org.apache.poi.xddf.usermodel.chart.XDDFChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFChartLegend;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSourcesFactory;
import org.apache.poi.xddf.usermodel.chart.XDDFLineChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFNumericalDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFValueAxis;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Component;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.text.DateFormatSymbols;
import java.util.Arrays;
import java.util.Iterator;
import java.util.List;
import java.util.Random;
import java.util.stream.IntStream;

@Component
public class ExcelUtils {


    public void createExcelFile() {
        String fileName = "test.xlsx";
        String chartLegend = "Teste";

        try (
                XSSFWorkbook xssfWorkbook = new XSSFWorkbook();
                OutputStream outputStream = new FileOutputStream(fileName)
        ) {

            XSSFSheet lineChartSheetTab = xssfWorkbook.createSheet("lineChartSheet");

            CellStyle cellCenterStyle = this.createCenterCellStyle(xssfWorkbook);
            CellStyle cellJustifyStyle = this.createJustifyCellStyle(xssfWorkbook);

            List<String> headers = Arrays.asList("A very long par   am1 name", "Param2", "Param3", "Param4", "Param5");
            List<String> months = Arrays.asList(new DateFormatSymbols().getMonths());

            this.createRowHeader(lineChartSheetTab, cellCenterStyle, headers);
            this.createColumnHeader(lineChartSheetTab, cellJustifyStyle, months);

            // Max rows supported by Excel: 1.048.576 rows
            for (int i = 1; i < months.size(); i++) {
                Row row = lineChartSheetTab.getRow(i);
                for (int j = 0; j < headers.size(); j++) {
                    this.createCell(row, j + 1, cellJustifyStyle, this.populateRows());
                }
            }

            this.autoSizeColumns(xssfWorkbook);
            this.createChart(lineChartSheetTab, chartLegend);

            xssfWorkbook.write(outputStream);

        } catch (IOException e) {
            e.printStackTrace();
        }

    }

    private void createChart(XSSFSheet sheet, String legend) {
        int numRow = sheet.getPhysicalNumberOfRows();
        int numCol = sheet.getRow(sheet.getFirstRowNum()).getLastCellNum();

        XSSFDrawing shapes = sheet.createDrawingPatriarch();
        XSSFClientAnchor xssfClientAnchor = shapes.createAnchor(0, 0, 0, 0, 0, numRow + 1,
                12, numRow + 12);
        XSSFChart shapesChart = shapes.createChart(xssfClientAnchor);

        XDDFChartLegend xddfChartLegend = shapesChart.getOrAddLegend();
        xddfChartLegend.setPosition(LegendPosition.BOTTOM);

        shapesChart.setTitleText(legend);

        XDDFCategoryAxis bottomAxis = shapesChart.createCategoryAxis(AxisPosition.BOTTOM);
        bottomAxis.setTitle("Months");
        bottomAxis.setCrosses(AxisCrosses.AUTO_ZERO);

        XDDFValueAxis leftAxis = shapesChart.createValueAxis(AxisPosition.LEFT);
        leftAxis.setTitle("Params");
        leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);

        this.plotLineChart(sheet, shapesChart, bottomAxis, leftAxis, numCol);
    }

    private void plotLineChart(XSSFSheet sheet, XSSFChart shapesChart, XDDFCategoryAxis bottomAxis,
                               XDDFValueAxis leftAxis, int numCol) {

        XDDFLineChartData xddfLineChartData = (XDDFLineChartData) shapesChart.createData(ChartTypes.LINE, bottomAxis, leftAxis);

        XDDFDataSource<String> months = XDDFDataSourcesFactory.fromStringCellRange(sheet, new CellRangeAddress(1, 12, 0, 0));

        for (int i = 1; i < numCol; i++) {
            XDDFNumericalDataSource<Double> params = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(1, 12, i, i));
            XDDFChartData.Series addSeries = xddfLineChartData.addSeries(months, params);
            CellReference titleSeries = new CellReference(sheet.getSheetName(), 0, i, false, false);
            addSeries.setTitle(null, titleSeries);
        }

        xddfLineChartData.setVaryColors(false);

        shapesChart.plot(xddfLineChartData);
    }


    private Integer populateRows() {
        return new Random().nextInt(100000);
    }

    private void createRowHeader(XSSFSheet xssfSheet, CellStyle style, List<String> headerNames) {

        Row row = xssfSheet.createRow(0);

        IntStream.range(0, headerNames.size())
                .forEach(nameIndex -> this.createCell(row, nameIndex + 1, style, headerNames.get(nameIndex)));
    }

    private void createColumnHeader(XSSFSheet xssfSheet, CellStyle style, List<String> headerNames) {
        IntStream.range(0, headerNames.size())
                .forEach(nameIndex -> createCell(xssfSheet.createRow(nameIndex + 1), 0, style, headerNames.get(nameIndex)));
    }

    private <K> void createCell(Row row, int column, CellStyle style, K value) {
        Cell cell = row.createCell(column);
        cell.setCellStyle(style);

        if (value instanceof String) {
            cell.setCellValue((String) value);
        } else if (value instanceof Number) {
            cell.setCellValue(new BigDecimal(Double.toString(((Number) value).doubleValue())).doubleValue());
        } else {
            throw new RuntimeException("Incompatible type of value");
        }
    }

    private CellStyle createCenterCellStyle(Workbook workbook) {
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        return cellStyle;
    }

    private CellStyle createJustifyCellStyle(Workbook workbook) {
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setAlignment(HorizontalAlignment.LEFT);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        return cellStyle;
    }

    private void autoSizeColumns(XSSFWorkbook workbook) {
        int numberOfSheets = workbook.getNumberOfSheets();
        for (int i = 0; i < numberOfSheets; i++) {
            Sheet sheet = workbook.getSheetAt(i);
            if (sheet.getPhysicalNumberOfRows() > 0) {
                Row row = sheet.getRow(sheet.getFirstRowNum());
                Iterator<Cell> cellIterator = row.cellIterator();
                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    int columnIndex = cell.getColumnIndex();
                    sheet.autoSizeColumn(columnIndex);
                    int currentColumnWidth = sheet.getColumnWidth(columnIndex);
                    sheet.setColumnWidth(columnIndex, (currentColumnWidth + 2500));
                }

            }
        }
    }

}
