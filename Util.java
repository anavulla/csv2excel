

 @SneakyThrows
    public static String csv2excel(String csvFile) {
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");

        String fileNameWithOutExt = FilenameUtils.removeExtension(csvFile);
        String outputFile = fileNameWithOutExt + ".xlsx";
        LOGGER.info("converting {} to Excel {} ", csvFile, outputFile);

        ArrayList arList = null;
        ArrayList al = null;

        String thisLine;
        int count = 0;
        FileInputStream fis = new FileInputStream(csvFile);
        DataInputStream myInput = new DataInputStream(fis);
        int i = 0;
        arList = new ArrayList();
        while ((thisLine = myInput.readLine()) != null) {
            al = new ArrayList();
            String[] strar = thisLine.split(",");
            Collections.addAll(al, strar);
            arList.add(al);
            System.out.println();
            i++;
        }

        XSSFWorkbook workbook = new XSSFWorkbook();
        CreationHelper createHelper = workbook.getCreationHelper();
        XSSFSheet sheet = workbook.createSheet("sheet");
        CellStyle cellStyle = workbook.createCellStyle();
        XSSFFont font = workbook.createFont();
        font.setFontHeight(11);
        cellStyle.setFont(font);

        for (int k = 0; k < arList.size(); k++) {
            ArrayList ardata = (ArrayList) arList.get(k);
            Row row = sheet.createRow(k);
            for (int p = 0; p < ardata.size(); p++) {
                Cell cell = row.createCell((short) p);
                cell.setCellStyle(cellStyle);

                String data = ardata.get(p).toString();

                if (data.startsWith("=")) {
                    data = data.replaceAll("\"", "");
                    data = data.replaceAll("=", "");
                } else if (data.startsWith("\"")) {
                    data = data.replaceAll("\"", "");
                } else {
                    data = data.replaceAll("\"", "");
                }

                if (data.contains("ADD_COMMA_HERE")) {
                    System.out.println("BEFORE->" + data);
                    // handling commas in data
                    data = data.replaceAll("ADD_COMMA_HERE", ",");
                    System.out.println("AFTER--->" + data);
                }

                if (data.chars().anyMatch(Character::isLetter)) {
                    cell.setCellType(CellType.STRING);
                    cell.setCellValue(data);
                } else {
                    if (checkIfDateIsValid(data)) {
                        cellStyle = workbook.createCellStyle();
                        cellStyle.setDataFormat(createHelper.createDataFormat().getFormat("yyyy-MM-dd"));
                        Date dateValue = new SimpleDateFormat("yyyy-MM-dd", Locale.ENGLISH).parse(data);
                        cell.setCellValue(dateValue);
                        cell.setCellStyle(cellStyle);

                    } else {
                        cellStyle = workbook.createCellStyle();
                        cell.setCellStyle(cellStyle);
                        Double doubleValue = Doubles.tryParse(data);
                        if (null != doubleValue) {
                            cell.setCellValue(doubleValue);
                        } else {
                            cell.setCellValue(data);
                        }
                    }
                }
            }
        }

        FileOutputStream fileOut = new FileOutputStream(outputFile);
        workbook.write(fileOut);
        fileOut.close();

        return Paths.get(outputFile.trim()).toFile().getName();
    }
