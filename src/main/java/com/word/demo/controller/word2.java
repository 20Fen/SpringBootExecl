//package com.word.demo.controller;
//
///**
// * Description:
// *
// * @author yangfl
// * @date 2019年10月21日 10:11
// * Version 1.0
// */
//public class word2 {
//
//    static String fileName = "D:\\file\\新建 Microsoft Word 文档";
//
//    public static void main(String[] args) throws IOException {
//
//        testWord();
//
//    }
//
//    public static void testWord(){
//
//        ArrayList<User> list = new ArrayList<>();
//        list.clear();
//        try{
//            //载入文档最好格式为.doc后缀
//            //.docx后缀文件可能存在问题，可将.docx后缀文件另存为.doc
//            FileInputStream in = new FileInputStream(fileName+".doc");//载入文档
//            POIFSFileSystem pfs = new POIFSFileSystem(in);
//            HWPFDocument hwpf = new HWPFDocument(pfs);
//            Range range = hwpf.getRange();//得到文档的读取范围
//            TableIterator it = new TableIterator(range);
//
//            //迭代文档中的表格
//            while (it.hasNext()) {
//
//
//                Table tb = (Table) it.next();
//                // 原本在此处新建 对象User user = new User();
//                // 但导出的数量不对
//                //迭代行，默认从0开始
//                for (int i = 0; i < tb.numRows(); i++) {
//                    TableRow tr = tb.getRow(i);
//                    //迭代列，默认从0开始
//                    for (int j = 0; j < tr.numCells(); j++) {
//                        TableCell td = tr.getCell(j);//取得单元格
//                        //取得单元格的内容
//                        for(int k=0;k<td.numParagraphs();k++){
//
//
//                            Paragraph para =td.getParagraph(k);
//                            String s = para.text();
////		                      System.out.println(s);
//                            if(s.contains("点位名称")){
//                                User user = new User();
//                                list.add(user);
//                                list.get(list.size()-1).setName(td.getParagraph(k+1).text());
//                            }
//                            if(s.contains("所属区")){
//                                list.get(list.size()-1).setArea(td.getParagraph(k+1).text());
//                            }
//                            if(s.contains("维护日期")){
//                                list.get(list.size()-1).setDate(td.getParagraph(k+1).text());
//                            }
//                            if(s.contains("维护措施")){
//                                list.get(list.size()-1).setCuoshi(td.getParagraph(k+1).text());
//                            }
////		                      System.out.println(s);
//                        } //end for
//                    }   //end for
//                }   //end for
//            } //end while
//            createExecl(list);
//            System.out.println(list.size());
//        }catch(Exception e){
//            e.printStackTrace();
//        }
//    }// end method
//    public static void createExecl(ArrayList<User> list) {
//        // 第一步，创建一个webbook，对应一个Excel文件
//        HSSFWorkbook wb = new HSSFWorkbook();
//        // 第二步，在webbook中添加一个sheet,对应Excel文件中的sheet
//        HSSFSheet sheet = wb.createSheet("123465789");
//        // 第三步，在sheet中添加表头第0行,注意老版本poi对Excel的行数列数有限制short
//        HSSFRow row = sheet.createRow((int) 0);
//        // 第四步，创建单元格，并设置值表头 设置表头居中
//        HSSFCellStyle style = wb.createCellStyle();
////        style.setAlignment(HSSFCellStyle.ALIGN_CENTER); // 创建一个居中格式
//        HSSFCell cell = row.createCell((short) 0);
//        cell.setCellValue("点位名称");
//        cell.setCellStyle(style);
//        cell = row.createCell((short) 1);
//        cell.setCellValue("所属区");
//        cell.setCellStyle(style);
//        cell = row.createCell((short) 2);
//        cell.setCellValue("维护日期");
//        cell.setCellStyle(style);
//        cell = row.createCell((short) 3);
//        cell.setCellValue("维护措施");
//        cell.setCellStyle(style);
//        cell = row.createCell((short) 4);
//
//        for (int i = 0; i < list.size(); i++) {
//            row = sheet.createRow((int) i + 1);
//            User user = (User) list.get(i);
//            // 第四步，创建单元格，并设置值
//            row.createCell((short) 0).setCellValue(user.getName());
//            row.createCell((short) 1).setCellValue(user.getArea());
//            row.createCell((short) 2).setCellValue(user.getDate());
//            row.createCell((short) 3).setCellValue(user.getCuoshi());
//            // row.createCell((short) 2).setCellValue((double) stu.getAge());
//            // cell = row.createCell((short) 3);
//            // cell.setCellValue(new SimpleDateFormat("yyyy-mm-dd").format(stu
//            // .getBirth()));
//        }
//        // 第六步，将文件存到指定位置
//        try {
//            FileOutputStream fout = new FileOutputStream(fileName + ".xls");
//            // 选中项目右键，点击Refresh，即可显示导出文件
//            wb.write(fout);
//            fout.close();
//        } catch (Exception e) {
//            e.printStackTrace();
//        }
//
//    }
//
//}
//
//
//class User {
//
//    String area ;
//
//    String name;
//
//    String date;
//
//    String cuoshi;
//
//    public String getArea() {
//        return area;
//    }
//
//    public void setArea(String area) {
//        this.area = area;
//    }
//
//    public String getName() {
//        return name;
//    }
//
//    public void setName(String name) {
//        this.name = name;
//    }
//
//    public String getDate() {
//        return date;
//    }
//
//    public void setDate(String date) {
//        this.date = date;
//    }
//
//    public String getCuoshi() {
//        return cuoshi;
//    }
//
//    public void setCuoshi(String cuoshi) {
//        this.cuoshi = cuoshi;
//    }
//}
