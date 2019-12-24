package wu.excelOper;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.*;

public class Oper {
    private static String defPath = "D:\\Oper\\"; // 默认路径
    private static String splitDirPath = defPath + "splitDir\\"; // 总表文件位置

    private static int point = Integer.MAX_VALUE; // 需要根据拆分的字段
    private static String mergeDirPath = defPath; // 需要合并的目录名字


    // 在D盘根目录下创建与项目名相同的文件名
    static {
        // 创建工作目录
        // 创建成功，不执行下面内容，否则继续执行
        if (new File(splitDirPath).mkdirs()) {
            System.out.println("\t文件夹不存在，已创建：" + splitDirPath + "所有文件操作将在新建目录下");
//            System.exit(0);
        }
//        // 读取配置文件
//        InputStream is = null;
//        Properties pro = new Properties();
//        try {
//            is = Oper.class.getResourceAsStream("/filePro.properties");
//            pro.load(is);
//
//            String i = pro.getProperty("point");
//            if (i != null && !i.equals("")) {
//                point = Integer.parseInt(i);
//            }
//
//            mergeDirname = pro.getProperty("mergeDirname");
//        } catch (IOException e) {
//            e.printStackTrace();
//        } finally {
//            try {
//                if (is != null) {
//                    is.close();
//                }
//            } catch (IOException e) {
//                e.printStackTrace();
//            }
//        }
    }

    /**
     * 根据用户输入确切调用哪个函数
     */
    public static void cir() {
        // 命令行输入数据
        Scanner sc = new Scanner(System.in);

        while (true) {
            System.out.println("=================================================================");
            int i = 0; // 用户输入的选择

            System.out.println("\t\t如果要拆分文件，请输入1：");
            System.out.println("\t\t如果要合并文件，请输入2：");
            System.out.println("\t\t如果要退出程序，请输入3：");
            System.out.println("\t\t其他任意键将会重新输入");

            try {
                i = sc.nextInt();
                sc.nextLine();
            } catch (Exception e) {
                System.out.println("\t\t输入数据有误噢，数据已强制转化为 0");
            }

            switch (i) {
                // 拆分文件
                case 1:
                    // 指定拆分的字段
                    System.out.println("\t\t请输入所需要拆分的字段，下标从 0 开始");
                    moreSeparated(sc.nextLine());
                    break;

                //合并文件
                case 2:
                    System.out.println("\t\t请输入所需要合并的文件夹名称");
                    moreMerge(sc.nextLine());
                    break;

                // 退出程序
                case 3:
                    System.exit(0);
                    break;

                //错误处理
                default:
                    System.out.println("\t\t输入的数据为 0，有误噢，请重新输入");
            }


            System.out.println("=================================================================");
        }
    }

    /**
     * 批量拆分
     */
    private static void moreSeparated(String pointStr) {
        // 兼容性判断
        if (pointStr != null && !pointStr.equals("")) {
            try {
                point = Integer.parseInt(pointStr);
            } catch (Exception e) {
                System.out.println("\t\t异常，所填写的拆分的字段对应的下标可能有误");
            }
        } else {
            System.out.println("\t\t所填写的拆分的字段对应的下标可能有误");
        }

        // 遍历文件夹下所有的目录，全部拆分
        if (point != Integer.MAX_VALUE) {
            File[] files = new File(splitDirPath).listFiles();
            if (files != null) {
                if (files.length == 0) {
                    System.out.println("\t\t路径下目录为空，无法处理，请将所需要的拆分的文件加入到此目录：" + splitDirPath);
                }

                for (File file : files) {
                    separated(file.getAbsolutePath(), defPath, point);
                }
            } else {
                System.out.println("\t\tfiles 为空，程序内部可能出现错误");
            }
        }
    }

    /**
     * 批量合并
     */
    private static void moreMerge(String mergeDirnameStr) {
        mergeDirPath += mergeDirnameStr;

        // 兼容性判断
        File tempFile = new File(mergeDirPath);
        if (!tempFile.exists()) {
            System.out.println("\t\t文件路径不存在：" + mergeDirPath);
            return;
        }

        File[] tempFiles = tempFile.listFiles();
        if (tempFiles != null && tempFiles.length == 0) {
            System.out.println("\t\t文件路径对应文件夹为空，无法处理，请将所需要的拆分的文件加入到此目录：" + mergeDirPath);
            return;
        }

        // 遍历文件夹下所有的目录，全部合并
        if (mergeDirnameStr != null && !mergeDirnameStr.equals("")) {
            merge(mergeDirPath, defPath);
        }
    }

    /**
     * 将文件拆分
     *
     * @param InPath  文件路径
     * @param outPath 拆分文件的位置
     * @param point   需要根据拆分的字段
     */
    private static void separated(String InPath, String outPath, int point) {
        File file = new File(InPath);
        ReadExcel readExcel = new ReadExcel();
        List<String[]> dataList = readExcel.readExcel(file);

        // 移除，获取属性头
        String[] head = dataList.remove(0);

        // 排序
        dataList.sort(new Comparator<String[]>() {
            @Override
            public int compare(String[] o1, String[] o2) {
                return o1[point].compareTo(o2[point]);
            }
        });

        // 分组
        // 找到第一个排序元素
        String s = (String) dataList.get(0)[point];
        // 如果获取到，说明存在
        while (s != null) {
            // 存放单个文件的数据
            List<String[]> data = new ArrayList<String[]>();
            // 将属性头加入
            data.add(head);

            // 查找，如果字段存在，加入并删除
            while (dataList.size() > 0 && s.equals(dataList.get(0)[point])) {
                data.add(dataList.remove(0));
            }

            // 创建输出位置
            File outfile = new File(outPath, s);
            if (!outfile.exists()) {
                boolean mkdirs = outfile.mkdirs();
            }

            // 输出文件
            WriteExcel writeExcel = new WriteExcel();
            // 输出到对应的文件夹
            String outfileName = file.getName().substring(0, file.getName().lastIndexOf("."));
            writeExcel.writeExcel(data, outPath + s + "\\" + outfileName + "_" + s + ".xlsx");

            // 获取新分组if
            if (dataList.size() > 0) {
                s = (String) dataList.get(0)[point];
            } else {
                s = null;
            }
        }
    }

    /**
     * 将文件合并
     *
     * @param InPath  分离文件存放路径
     * @param outPath 输出文件路径
     */
    private static void merge(String InPath, String outPath) {
        File path = new File(InPath);
        File[] files = path.listFiles();

        // 总数据
        List<String[]> allData = new ArrayList<>();
        String[] head = null;

        // 将所有数据读入内存中
        if (files != null) {
            for (File file : files) {
                ReadExcel ReadExcel = new ReadExcel();
                List<String[]> datalist = ReadExcel.readExcel(file);

                // 删除属性
                head = datalist.remove(0);

                // 将数据加入到新表中
                allData.addAll(datalist);
            }

            // 将属性头插入
            allData.add(0, head);

            WriteExcel writeExcel = new WriteExcel();
            writeExcel.writeExcel(allData, outPath + "all_" + path.getName() + ".xlsx");
        }
    }
}
