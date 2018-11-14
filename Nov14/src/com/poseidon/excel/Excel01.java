package com.poseidon.excel;

import java.io.File;
import java.util.ArrayList;
import java.util.HashMap;

import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

public class Excel01 {
	public static void main(String[] args) {
		WritableWorkbook workbook = null;

		WritableSheet sheet = null;
		Label label = null;
		File file = new File("c:\\temp\\excel02.xls");

		HashMap<String, String> hashMap = new HashMap<String, String>();
		hashMap.put("�̸�", "ȫ�浿");
		hashMap.put("����", "30");
		hashMap.put("����", "=NOW()- YEAR(1950)");

		HashMap<String, String> hashMap2 = new HashMap<>();
		hashMap2.put("�̸�", "��浿");
		hashMap2.put("����", "20");
		hashMap2.put("����", "=NOW()- YEAR(1950)");

		ArrayList<HashMap<String, String>> arrayList = new ArrayList<>();
		arrayList.add(hashMap);
		arrayList.add(hashMap2);

		try {
			workbook = Workbook.createWorkbook(file);

			workbook.createSheet("��Ʈ", 0);
			sheet = workbook.getSheet(0);

			label = new Label(0, 0, "�̸�");
			sheet.addCell(label);

			label = new Label(1, 0, "����");
			sheet.addCell(label);

			label = new Label(2, 0, "����");
			sheet.addCell(label);

			for (int i = 0; i < arrayList.size(); i++) {
				HashMap<String, String> rs = arrayList.get(i);

				label = new Label(0, (i + 1), rs.get("�̸�"));
				sheet.addCell(label);

				label = new Label(1, (i + 1), rs.get("����"));
				sheet.addCell(label);
				
				label = new Label(2, (i + 1), rs.get("����"));
				sheet.addCell(label);
			}
			workbook.write();
			workbook.close();

		} catch (Exception e) {
			e.printStackTrace();
		}

	}
}
