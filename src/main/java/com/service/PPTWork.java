package com.service;
import java.io.FileInputStream;

import org.apache.poi.xslf.usermodel.XMLSlideShow;


import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Iterator;
import java.util.List;
import java.util.ResourceBundle;

import org.apache.poi.sl.usermodel.PictureData;
import org.apache.poi.sl.usermodel.PictureData.PictureType;
import org.apache.poi.util.IOUtils;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFGroupShape;
import org.apache.poi.xslf.usermodel.XSLFPictureData;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTable;
import org.apache.poi.xslf.usermodel.XSLFTableCell;
import org.apache.poi.xslf.usermodel.XSLFTableRow;
import org.apache.poi.xslf.usermodel.XSLFTextBox;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.json.JSONArray;
import org.json.JSONObject;

public class PPTWork {

	ResourceBundle bundleststic = ResourceBundle.getBundle("config_PPTExcel");

	public static void main(String[] args) {
		// TODO Auto-generated method stub

	}

	public String parsePPT(String ppatpath, JSONObject datajson, String savepath) {
		
		JSONObject imagesobject=new JSONObject();
		JSONArray imagesArray=null;
		
		try {
			
			if(datajson.has("imagesArray")) {
				imagesArray=datajson.getJSONArray("imagesArray");
				for(int i=0; i<imagesArray.length(); i++) {
					JSONObject subobj=imagesArray.getJSONObject(i);
					String fieldname=subobj.getString("fieldname");
					String fieldvalue=subobj.getString("fieldvalue");
					imagesobject.put(fieldname, fieldvalue)	;
				}
			}
			System.out.println("imagesobject: " + imagesobject);

			XMLSlideShow ppt = new XMLSlideShow(new FileInputStream(ppatpath));
			for (XSLFSlide slide : ppt.getSlides()) {
				System.out.println("REDACT Slide: **********************************************************" );

				System.out.println("REDACT Slide: " + slide.getTitle());

				XSLFTextShape[] Textshapes = slide.getPlaceholders();
				System.out.println("Textshapes.length "+Textshapes.length);
				for (XSLFTextShape textShape : Textshapes) {
					System.out.println("textShape name "+textShape.getShapeName());

					System.out.println("textShape.getShapeName() "+ textShape.getShapeName());
					System.out.println("textShape.getTextType() "+ textShape.getTextType());

					List<XSLFTextParagraph> textparagraphs = textShape.getTextParagraphs();
					for (XSLFTextParagraph para : textparagraphs) {
						List<XSLFTextRun> textruns = para.getTextRuns();
						for (XSLFTextRun incomingTextRun : textruns) {
							String text = incomingTextRun.getRawText();
							System.out.println(text);

							// first check for image add
							if (text.indexOf("{{") != -1) {
								if (imagesobject.has(text)) {
									String imagename = "";
									String imagesavepath = "";
									String imagelink = "";
									String extension = "";

									if (imagelink.lastIndexOf("/") != -1) {
										imagename = imagelink.substring(imagelink.lastIndexOf("/") + 1);
									}
									if (imagelink.lastIndexOf(".") != -1) {
										extension = imagelink.substring(imagelink.lastIndexOf("") + 1);
									}

									imagelink = imagesobject.getString(text);
									imagesavepath = bundleststic.getString("images_path");
									System.out.println("imagename "+imagename);
									System.out.println("imagesavepath "+imagesavepath);
									System.out.println("imagelink "+imagelink);
									System.out.println("extension  "+extension);
									

									String resp = new SOAPCall().saveTemplate(imagelink, imagesavepath, imagename);
									if (resp.equalsIgnoreCase("success")) {
										System.out.println("image loation " + imagesavepath + imagename);
										File image = new File(imagesavepath + imagename);
										byte[] picture = IOUtils.toByteArray(new FileInputStream(image));
										PictureData idx = null;
										if (extension.equalsIgnoreCase("png")) {
											idx = ppt.addPicture(picture, PictureType.PNG);
										} else if (extension.equalsIgnoreCase("jpeg")) {
											idx = ppt.addPicture(picture, PictureType.JPEG);
										}else {
											idx = ppt.addPicture(picture, PictureType.PNG);
										}
										XSLFPictureShape pic = slide.createPicture(idx);
										pic.setAnchor(textShape.getAnchor());
									}
								}
							} else if (text.indexOf("{{") == -1 && datajson.has(text)) {
								String value = datajson.getString(text);
								System.out.println("value" + value);
								String newText = text.replaceAll("(?i)" + text, value);
								System.out.println("final text 1 " + newText);

								incomingTextRun.setText(newText);
							} else {
								System.out.println("in else");

								while (text.indexOf("<<") != -1 && text.indexOf(">>") != -1) {
									System.out.println("textelementstr.indexOf(\"<<\"): " + text.indexOf("<<"));
									System.out.println("textelementstr.indexOf(\">>\"): " + text.indexOf(">>"));

									String key = text.substring(text.indexOf("<<"), text.indexOf(">>") + 2);
									System.out.println("key " + key + " textelementstr " + text);
									if (datajson.has(key)) {
										System.out.println("paramsMap.get(key)" + datajson.getString(key).toString());
										text = text.replace(key, (String) datajson.getString(key));
									} else {
										text = text.replace(key, "");
									}
								}

								while (text.indexOf("<$") != -1 && text.indexOf("$>") != -1) {
									System.out.println("textelementstr.indexOf(\"<$\"): " + text.indexOf("<$"));
									System.out.println("textelementstr.indexOf(\"$>\"): " + text.indexOf("$>"));
									String key = text.substring(text.indexOf("<$"), text.indexOf("$>") + 2);
									System.out.println("key " + key + " textelementstr " + text);
									if (datajson.has(key)) {
										text = text.replace(key, (String) datajson.getString(key));
										System.out.println("key " + key + " textelementstr " + text);
									} else {
										text = text.replace(key, "");
									}
								}

								while (text.indexOf("$<{") != -1 && text.indexOf("}>$") != -1) {
									System.out.println("textelementstr.indexOf(\"$<{\"): " + text.indexOf("$<{"));
									System.out.println("textelementstr.indexOf(\"}>$\"): " + text.indexOf("}>$"));

									String key = text.substring(text.indexOf("$<{"), text.indexOf("}>$") + 2);
									System.out.println("key " + key + " textelementstr " + text);
									if (datajson.has(key)) {
										text = text.replace(key, (String) datajson.getString(key));
										System.out.println("key " + key + " textelementstr " + text);
									} else {
										text = text.replace(key, "");
									}
								}
								while (text.indexOf("${") != -1 && text.indexOf("}$") != -1) {
									System.out.println("textelementstr.indexOf(\"${\"): " + text.indexOf("${"));
									System.out.println("textelementstr.indexOf(\"}$\"): " + text.indexOf("}$"));

									String key = text.substring(text.indexOf("${"), text.indexOf("}$") + 2);
									System.out.println("key " + key + " textelementstr " + text);
									if (datajson.has(key)) {
										text = text.replace(key, datajson.getString(key));
										System.out.println("key " + key + " textelementstr " + text);
									} else {
										text = text.replace(key, "");
									}
								}
								System.out.println("final text2 " + text);
								incomingTextRun.setText(text);
							}
						}
					}
				}

				// parsing shapes start
				List<XSLFShape> groupshaps = slide.getShapes();
				System.out.println("slide table start =================================================================================================");

				for (XSLFShape tableShape : groupshaps) {
					if (tableShape instanceof XSLFTable) {

						System.out.println("table " + ((XSLFTable) tableShape).getCTTable());
						System.out.println("column " + ((XSLFTable) tableShape).getNumberOfColumns());
						System.out.println("rows " + ((XSLFTable) tableShape).getNumberOfRows());

//						XSLFTableRow addnewrow = ((XSLFTable) tableShape).addRow();
//						for (int c = 0; c < 4; c++) {
//							XSLFTextRun newaddcell = addnewrow.addCell().setText("cell no " + c);
//							System.out.println("cell added");
//						}

						Iterator<XSLFTableRow> itr = ((XSLFTable) tableShape).iterator();
						while (itr.hasNext()) {
							System.out.println("****new row *****");
							XSLFTableRow row = itr.next();
							List<XSLFTableCell> cells = row.getCells();
							for (XSLFTableCell cell : cells) {
								// cell.getTextParagraphs()
								List<XSLFTextParagraph> textparagraphs = cell.getTextParagraphs();
								for (XSLFTextParagraph para : textparagraphs) {
									List<XSLFTextRun> textruns = para.getTextRuns();
									for (XSLFTextRun incomingTextRun : textruns) {
										String text = incomingTextRun.getRawText();
										System.out.println(text);
										if ( datajson.has(text)) {
											String value = datajson.getString(text);
											System.out.println("value" + value);
											String newText = text.replaceAll("(?i)" + text, value);
											System.out.println("final text 1 " + newText);
											incomingTextRun.setText(newText);
										} else {
											System.out.println("in else");
											while (text.indexOf("<<") != -1 && text.indexOf(">>") != -1) {
												System.out.println("textelementstr.indexOf(\"<<\"): " + text.indexOf("<<"));
												System.out.println("textelementstr.indexOf(\">>\"): " + text.indexOf(">>"));

												String key = text.substring(text.indexOf("<<"), text.indexOf(">>") + 2);
												System.out.println("key " + key + " textelementstr " + text);
												if (datajson.has(key)) {
													System.out.println("datajson.get(key)" + datajson.getString(key).toString());
													text = text.replace(key, (String) datajson.getString(key));
												} else {
													text = text.replace(key, "");
												}
											}
											incomingTextRun.setText(text);
										}
									}
								}
							}

						}
					}
					
					System.out.println("slide table end =================================================================================================");
					System.out.println("slide textbox start =================================================================================================");

					if(tableShape instanceof XSLFTextBox ) {
					List<XSLFTextParagraph> textparagraphs = ((XSLFTextBox) tableShape).getTextParagraphs();
					for (XSLFTextParagraph para : textparagraphs) {

					List<XSLFTextRun> textruns = para.getTextRuns();

					for (XSLFTextRun incomingTextRun : textruns) {

					String text = incomingTextRun.getRawText();

					System.out.println(slide.getSlideNumber() + "textparagraphs :: "+text);
					if ( datajson.has(text)) {
						String value = datajson.getString(text);
						System.out.println("value" + value);
						String newText = text.replaceAll("(?i)" + text, value);
						System.out.println("final text 1 " + newText);

						incomingTextRun.setText(newText);
					} else {
						System.out.println("in else");

						while (text.indexOf("<<") != -1 && text.indexOf(">>") != -1) {
							System.out.println("textelementstr.indexOf(\"<<\"): " + text.indexOf("<<"));
							System.out.println("textelementstr.indexOf(\">>\"): " + text.indexOf(">>"));

							String key = text.substring(text.indexOf("<<"), text.indexOf(">>") + 2);
							System.out.println("key " + key + " textelementstr " + text);
							if (datajson.has(key)) {
								System.out.println("paramsMap.get(key)" + datajson.getString(key).toString());
								text = text.replace(key, (String) datajson.getString(key));
							} else {
								text = text.replace(key, "");
							}
						}
						incomingTextRun.setText(text);

					}
				
					}
					}
					}	
					System.out.println("slide textbox end =================================================================================================");

					
					
				}
			}
System.out.println("savepath "+savepath);
			  FileOutputStream out = new FileOutputStream(savepath);
			  ppt.write(out);
			 // slideShow.close();
			  out.close();

		} catch (Exception e) {
			e.printStackTrace();
		}

		return null;
	}

}
