package com.demo.writeExcel;

import java.util.Date;

public class Invoices {
	
	Integer itemID;
	String itemName;
	Integer itemQty;
	Double totalPrice;
	Date itemSoldDate;
	public Invoices(Integer itemID, String itemName, Integer itemQty, Double totalPrice, Date itemSoldDate) {
		super();
		this.itemID = itemID;
		this.itemName = itemName;
		this.itemQty = itemQty;
		this.totalPrice = totalPrice;
		this.itemSoldDate = itemSoldDate;
	}
	public Integer getItemID() {
		return itemID;
	}
	public void setItemID(Integer itemID) {
		this.itemID = itemID;
	}
	public String getItemName() {
		return itemName;
	}
	public void setItemName(String itemName) {
		this.itemName = itemName;
	}
	public Integer getItemQty() {
		return itemQty;
	}
	public void setItemQty(Integer itemQty) {
		this.itemQty = itemQty;
	}
	public Double getTotalPrice() {
		return totalPrice;
	}
	public void setTotalPrice(Double totalPrice) {
		this.totalPrice = totalPrice;
	}
	public Date getItemSoldDate() {
		return itemSoldDate;
	}
	public void setItemSoldDate(Date itemSoldDate) {
		this.itemSoldDate = itemSoldDate;
	}
	

}
