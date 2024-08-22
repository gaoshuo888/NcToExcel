package excelsummary;

/**
 * FileName: ${NAME}.java
 * 类的详细说明
 *
 * @author GaoShuo
 * @version 1.0.0
 * @Date 2024/8/22
 */
public class Main {
    public static void main(String[] args) {
        //需要汇总的文件路径及输出文件的路径
        String inputFile = "E:\\DownLoad\\TotalP\\output2.xlsx";
        String outputFile = "E:\\DownLoad\\TotalP\\01.xlsx";
        //需要汇总的单元格
        String[] cellNames = {"B2", "C2", "B3"};
        //时间格式转换，最终文件
        String inputFilePath = outputFile;
        String outputFilePath = "E:/DownLoad/TotalP/final.xlsx";

        DataSummary DataSummary =new DataSummary();
        DataSummary.setInputFile(inputFile);
        DataSummary.setOutputFile(outputFile);
        DataSummary.setCellNames(cellNames);
        DataSummary.dataSummary();

        ReviseTimeFormat ReviseTimeFormat = new ReviseTimeFormat(inputFilePath , outputFilePath);
        ReviseTimeFormat.reviseTimeFormat();

    }
}