package main

import (
	"database/sql"
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
	"github.com/gin-gonic/gin"
	_ "github.com/lib/pq"
	"log"
	"net/http"
	"sort"
	"strconv"
	"time"
	// "github.com/robfig/cron/v3"
)

type Transaction struct {
	id                 int
	date               string
	merchant_id         string
	shop_id            string
	shop_name           string
	order_no            string
	transaction_type    string
	description        string
	order_time          string
	order_completed_time string
	amount             int
	promo_m             int
	promo_ltj           int
	gp                 int
	vat_on_gp            float32
	wht                float32
	net_food_amount      int
	payment_method      string
	settlement_time     string
}

func exportToExcel(db *sql.DB, c *gin.Context) {

	uniqueMerchantID, err := db.Query("SELECT DISTINCT merchant_id FROM transactions")
	if err != nil {
		c.JSON(http.StatusInternalServerError, gin.H{"error": err.Error()})
		return
	}
	defer uniqueMerchantID.Close()
	t := time.Now()

	var merchantIDs []string // Use the appropriate type for merchant_id

	for uniqueMerchantID.Next() {
		var merchantID string // Use the appropriate type for merchant_id
		if err := uniqueMerchantID.Scan(&merchantID); err != nil {
			c.JSON(http.StatusInternalServerError, gin.H{"error": err.Error()})
			return
		}
		merchantIDs = append(merchantIDs, merchantID)
	}
	sort.Strings(merchantIDs)

	f := excelize.NewFile()
	sheetName := "Sheet1"

	// Styels
	yellowFillStyle, err := f.NewStyle(`{
		"fill": {
		"type": "pattern",
		"color": ["#FFFF00"],
		"pattern": 1
		},
		"border": [
			{
				"type": "left",
				"color": "000000",
				"style": 1
			},
			{
				"type": "right",
				"color": "000000",
				"style": 1
			},
			{
				"type": "top",
				"color": "000000",
				"style": 1
			},
			{
				"type": "bottom",
				"color": "000000",
				"style": 1
			}
		],
		"alignment": {
			"horizontal": "center",
			"vertical": "center"
		},
		"padding": {
			"left": 5,
			"right": 5
		}
	}`)
	if err != nil {
		c.JSON(500, gin.H{"error": err.Error()})
		return
	}

	redFontStyle, err := f.NewStyle(`{
	"font":{"color":"#FF0000"},
	"border": [
			{
				"type": "left",
				"color": "000000",
				"style": 1
			},
			{
				"type": "right",
				"color": "000000",
				"style": 1
			},
			{
				"type": "top",
				"color": "000000",
				"style": 1
			},
			{
				"type": "bottom",
				"color": "000000",
				"style": 1
			}
		],
		"alignment": {
			"horizontal": "center",
			"vertical": "center"
		},
		"padding": {
			"left": 5,
			"right": 5
		}
	}`)
	if err != nil {
		c.JSON(500, gin.H{"error": err.Error()})
		return
	}
	subTotalStyle, err := f.NewStyle(`{
		"font":{"color":"#FF0000"},
		"border": [
				{
					"type": "left",
					"color": "000000",
					"style": 2
				},
				{
					"type": "right",
					"color": "000000",
					"style": 2
				},
				{
					"type": "top",
					"color": "000000",
					"style": 2
				},
				{
					"type": "bottom",
					"color": "000000",
					"style": 2
				}
			],
			"alignment": {
				"horizontal": "center",
				"vertical": "center"
			},
			"padding": {
				"left": 10,
				"right": 10
			}
		}`)
	if err != nil {
		c.JSON(500, gin.H{"error": err.Error()})
		return
	}
	blankStyle, err := f.NewStyle(`{
		"border": [
			{
				"type": "left",
				"color": "000000",
				"style": 1
			},
			{
				"type": "right",
				"color": "000000",
				"style": 1
			},
			{
				"type": "top",
				"color": "000000",
				"style": 1
			},
			{
				"type": "bottom",
				"color": "000000",
				"style": 1
			}
		],
		"alignment": {
			"horizontal": "center",
			"vertical": "center"
		},
		"padding": {
			"left": 5,
			"right": 5
		}
	}`)
	if err != nil {
		c.JSON(500, gin.H{"error": err.Error()})
		return
	}
	fontRed_bgYellow, err := f.NewStyle(`{
		"fill": {
		"type": "pattern",
		"color": ["#FFFF00"],
		"pattern": 1
		},
		"font":{"color":"#FF0000"},
		
		"border": [
			{
				"type": "left",
				"color": "000000",
				"style": 1
			},
			{
				"type": "right",
				"color": "000000",
				"style": 1
			},
			{
				"type": "top",
				"color": "000000",
				"style": 1
			},
			{
				"type": "bottom",
				"color": "000000",
				"style": 1
			}
		],
		"alignment": {
			"horizontal": "center",
			"vertical": "center"
		},
		"padding": {
			"left": 5,
			"right": 5
		}
	}`)
	if err != nil {
		c.JSON(500, gin.H{"error": err.Error()})
		return
	}


	rows, err := db.Query("SELECT * FROM transactions")
	if err != nil {
		c.JSON(http.StatusInternalServerError, gin.H{"error": err.Error()})
		return
	}
	defer rows.Close()
	var transactions [] Transaction
	for rows.Next() {
		var txn Transaction
		err := rows.Scan(&txn.id, &txn.date, &txn.merchant_id, &txn.shop_id, &txn.shop_name, &txn.order_no, &txn.transaction_type,
			&txn.description, &txn.order_time, &txn.order_completed_time, &txn.amount, &txn.promo_m, &txn.promo_ltj,
			&txn.gp, &txn.vat_on_gp, &txn.wht, &txn.net_food_amount, &txn.payment_method, &txn.settlement_time)
	
		if err != nil {
			c.JSON(500, gin.H{"error": err.Error()})
			return
		}
		
		transactions = append(transactions, txn)
	}
	f.SetColWidth(sheetName, "A", "A", 30.00)
	f.SetColWidth(sheetName, "B", "C", 20.00)
	f.SetColWidth(sheetName, "E", "E", 30.00)
	f.SetColWidth(sheetName, "F", "F", 20.00)
	f.SetColWidth(sheetName, "G", "G", 25.00)
	f.SetColWidth(sheetName, "H", "I", 30.00)
	f.SetColWidth(sheetName, "J", "N", 15.00)
	f.SetColWidth(sheetName, "O", "P", 20.00)
	f.SetColWidth(sheetName, "Q", "Q", 15.00)
	f.SetColWidth(sheetName, "R", "R", 30.00)

	log_count := 1
	numCheckHeader := 1
	for _, unique_merchantID_id := range merchantIDs {
		fmt.Println(unique_merchantID_id)
		f.SetCellValue(sheetName, "A"+strconv.Itoa(numCheckHeader), "Date")
		f.SetCellValue(sheetName, "B"+strconv.Itoa(numCheckHeader), "Merchant ID")
		f.SetCellValue(sheetName, "C"+strconv.Itoa(numCheckHeader), "Shop ID")
		f.SetCellValue(sheetName, "D"+strconv.Itoa(numCheckHeader), "Shop Name")
		f.SetCellValue(sheetName, "E"+strconv.Itoa(numCheckHeader), "Order No")
		f.SetCellValue(sheetName, "F"+strconv.Itoa(numCheckHeader), "Transaction Type")
		f.SetCellValue(sheetName, "G"+strconv.Itoa(numCheckHeader), "Description")
		f.SetCellValue(sheetName, "H"+strconv.Itoa(numCheckHeader), "Order Time")
		f.SetCellValue(sheetName, "I"+strconv.Itoa(numCheckHeader), "Order Completed Time")
		f.SetCellValue(sheetName, "J"+strconv.Itoa(numCheckHeader), "จำนวนเงิน")
		f.SetCellValue(sheetName, "K"+strconv.Itoa(numCheckHeader), "Promo M")
		f.SetCellValue(sheetName, "L"+strconv.Itoa(numCheckHeader), "Promo.LTJ")
		f.SetCellValue(sheetName, "M"+strconv.Itoa(numCheckHeader), "GP")
		f.SetCellValue(sheetName, "N"+strconv.Itoa(numCheckHeader), "VAT on GP")
		f.SetCellValue(sheetName, "O"+strconv.Itoa(numCheckHeader), "WHT")
		f.SetCellValue(sheetName, "P"+strconv.Itoa(numCheckHeader), "ค่าอาหารสุทธิ")
		f.SetCellValue(sheetName, "Q"+strconv.Itoa(numCheckHeader), "วิธีการชำระเงิน")
		f.SetCellValue(sheetName, "R"+strconv.Itoa(numCheckHeader), "Settlement Time")

		f.SetCellStyle(sheetName, "A1", "O1", blankStyle)
		f.SetCellStyle(sheetName, "P1", "P1", yellowFillStyle)
		f.SetCellStyle(sheetName, "R1", "S1", blankStyle)
		
		

	
		discard := 0
		subTotal := 0
		rowNumber := numCheckHeader + 1
		for _ , i :=  range transactions  {

			if err != nil {
				c.JSON(500, gin.H{"error": err.Error()})
				return
			}
			fmt.Printf("The sql is: %s\n", i.merchant_id)
			fmt.Printf("The merchant ID is: %s\n", unique_merchantID_id)

			fmt.Println(log_count)
			if i.merchant_id == unique_merchantID_id {
				f.SetCellValue(sheetName, "A"+strconv.Itoa(rowNumber), i.date)
				f.SetCellValue(sheetName, "B"+strconv.Itoa(rowNumber), i.merchant_id)
				f.SetCellValue(sheetName, "C"+strconv.Itoa(rowNumber), i.shop_id)
				f.SetCellValue(sheetName, "D"+strconv.Itoa(rowNumber), i.shop_name)
				f.SetCellValue(sheetName, "E"+strconv.Itoa(rowNumber), i.order_no)
				f.SetCellValue(sheetName, "F"+strconv.Itoa(rowNumber), i.transaction_type)
				f.SetCellValue(sheetName, "G"+strconv.Itoa(rowNumber), i.description)
				f.SetCellValue(sheetName, "H"+strconv.Itoa(rowNumber), i.order_time)
				f.SetCellValue(sheetName, "I"+strconv.Itoa(rowNumber), i.order_completed_time)
				f.SetCellValue(sheetName, "J"+strconv.Itoa(rowNumber), i.amount)
				f.SetCellValue(sheetName, "K"+strconv.Itoa(rowNumber), i.promo_m)
				f.SetCellValue(sheetName, "L"+strconv.Itoa(rowNumber), i.promo_ltj)
				f.SetCellValue(sheetName, "M"+strconv.Itoa(rowNumber), i.gp)
				f.SetCellValue(sheetName, "N"+strconv.Itoa(rowNumber), i.vat_on_gp)
				f.SetCellValue(sheetName, "O"+strconv.Itoa(rowNumber), i.wht)
				f.SetCellValue(sheetName, "P"+strconv.Itoa(rowNumber), i.net_food_amount)
				f.SetCellValue(sheetName, "Q"+strconv.Itoa(rowNumber), i.payment_method)
				f.SetCellValue(sheetName, "R"+strconv.Itoa(rowNumber), i.settlement_time)

				if i.transaction_type == "Void" || i.transaction_type == "Adjust" || i.transaction_type == "Refund" {
					discard += i.net_food_amount
					f.SetCellStyle(sheetName, "A"+strconv.Itoa(rowNumber), "O"+strconv.Itoa(rowNumber), redFontStyle)
					f.SetCellStyle(sheetName, "P"+strconv.Itoa(rowNumber), "P"+strconv.Itoa(rowNumber), fontRed_bgYellow)
					f.SetCellStyle(sheetName, "Q"+strconv.Itoa(rowNumber), "R"+strconv.Itoa(rowNumber), redFontStyle)

				} else {
					subTotal += i.net_food_amount
					f.SetCellStyle(sheetName, "A"+strconv.Itoa(rowNumber), "O"+strconv.Itoa(rowNumber), blankStyle)
					f.SetCellStyle(sheetName, "P"+strconv.Itoa(rowNumber), "P"+strconv.Itoa(rowNumber), yellowFillStyle)
					f.SetCellStyle(sheetName, "Q"+strconv.Itoa(rowNumber), "R"+strconv.Itoa(rowNumber), blankStyle)
				}
				rowNumber++
				log_count++
			}
		}
		f.SetCellValue(sheetName, "O"+strconv.Itoa(rowNumber+1), "Sub Total = ")
		f.SetCellValue(sheetName, "P"+strconv.Itoa(rowNumber+1), subTotal-discard)
		f.SetCellStyle(sheetName, "O"+strconv.Itoa(rowNumber+1), "O"+strconv.Itoa(rowNumber+1), subTotalStyle)
		f.SetCellStyle(sheetName, "P"+strconv.Itoa(rowNumber+1), "P"+strconv.Itoa(rowNumber+1), subTotalStyle)
		numCheckHeader = rowNumber + 3

	}
	err = f.SaveAs("Data_"+t.Format("01-02-2006")+".xlsx")
	if err != nil {
		c.JSON(500, gin.H{"error": err.Error()})
		return
	}

	c.File("Data_"+t.Format("01-02-2006")+".xlsx")
}

func data(db *sql.DB, c *gin.Context) {
	rows, err := db.Query("SELECT * FROM transactions")

	if err != nil {
		c.JSON(http.StatusInternalServerError, gin.H{"error": err.Error()})
		return
	}
	defer rows.Close()

	var results []map[string]interface{}
	for rows.Next() {
		var id, amount, promo_ltj, promo_m, gp, net_food_amount int
		var vat_on_gp, wht float32
		var date, merchant_id, shop_id, shop_name, order_no,
			transaction_type, description, order_time, order_completed_time, payment_method, settlement_time string

		rows.Scan(&id, &date, &merchant_id, &shop_id, &shop_name, &order_no,
			&transaction_type, &description, &order_time, &order_completed_time, &amount,
			&promo_m, &promo_ltj, &gp, &vat_on_gp, &wht,
			&net_food_amount, &payment_method, &settlement_time)

		row := map[string]interface{}{
			"id":                   id,
			"date":                 date,
			"merchant_id":          merchant_id,
			"shop_id":              shop_id,
			"shop_name":            shop_name,
			"order_no":             order_no,
			"transaction_type":     transaction_type,
			"description":          description,
			"order_time":           order_time,
			"order_completed_time": order_completed_time,
			"amount":               amount,
			"promo_m":              promo_m,
			"promo_ltj":            promo_ltj,
			"gp":                   gp,
			"vat_on_gp":            vat_on_gp,
			"wht":                  wht,
			"net_food_amount":      net_food_amount,
			"payment_method":       payment_method,
			"settlement_time":      settlement_time,
		}
		results = append(results, row)
	}
	c.JSON(http.StatusOK, results)
}


func main() {
	connStr := "user=user password=password dbname=go_database host=localhost sslmode=disable"

	db, err := sql.Open("postgres", connStr)
	if err != nil {
		log.Fatal(err)
	}
	defer db.Close()

	router := gin.Default()
	router.GET("/export", func(c *gin.Context) {
		exportToExcel(db, c)
	})
	router.GET("/data", func(c *gin.Context) {
		data(db, c)
	})

	router.Run(":8080")
}
