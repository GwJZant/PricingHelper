SELECT tickets.BRAND AS [Brand], 
	   tickets.STYLE AS [Style],
	   MAX(tickets.LOOKUP) AS [Barcode_Lookup], 
	   tickets.STORE_ID AS [Store],
	   tickets.PRICE AS [Old_Retail], 
	   CAST(styles.MARGIN_PERCENT AS DECIMAL(18, 2)) AS [Old_Margin], 
	   tickets.DESCRIPTION AS [Description_1], 
	   styles.DESCRIPTION_2 AS [Description_2], 
	   tickets.DEPT AS [Department], 
	   tickets.TYP AS [Type],
	   tickets.OF1 AS [Season], 
	   tickets.OF5 AS [Promo],
	   SUM(buckets.QOH) AS [Available],
	   tickets.PRICE AS [Last_Price],
	   buckets.LAST_COST AS [Last_Cost],
	   ISNULL(parts.CONTACT_ID, 0) AS [Contact_Id], 
	   ISNULL((CASE contacts.COMPANY
			WHEN 'ZEIGLER`S DISTRIBUTIOR INC' THEN 'Zeig'
			WHEN 'PHILLIP`S PET SUPPLY' THEN 'Phil'
			WHEN 'NATURAL ANIMAL NUTRITION' THEN 'NAN'
			WHEN 'BRADLEY CALDWELL' THEN 'Brad'
			WHEN 'PET FOOD EXPERTS' THEN 'PFX'
			WHEN 'TICKNER`S INC.' THEN 'Tick'
			ELSE contacts.COMPANY
		END), tickets.BRAND) AS Vendor,
		ISNULL(parts.PART_NUM, 0) AS [Part_Num],
		ISNULL(parts.COST, 0) AS [Part_Cost],
		MAX(lookups.UPC) AS [UPC]
FROM VW_TICKETS tickets
INNER JOIN TB_STYLES styles
ON styles.STYLE_ID = tickets.STYLE_ID
INNER JOIN TB_SKU_BUCKETS buckets
ON buckets.SKU_ID = tickets.SKU_ID
AND buckets.SKU_BUCKET_ID = tickets.SKU_BUCKET_ID
AND buckets.STORE_ID = tickets.STORE_ID
INNER JOIN TB_PARTS parts
ON parts.STYLE_ID = tickets.STYLE_ID
INNER JOIN TB_CONTACTS contacts
ON contacts.CONTACT_ID = parts.CONTACT_ID
LEFT JOIN (
    SELECT SKU_ID, MAX(LOOKUP) AS [UPC]
    FROM TB_SKU_LOOKUPS
    WHERE TYP = 1
    GROUP BY SKU_ID
) lookups ON lookups.SKU_ID = tickets.SKU_ID
WHERE contacts.COMPANY = $(Vendor)
AND tickets.OF1 <> 'DISCO'
AND styles.STATUS_FINISH <> 'Y'
GROUP BY tickets.BRAND, 
	   tickets.STYLE,
	   tickets.STORE_ID,
	   tickets.PRICE, 
	   styles.MARGIN_PERCENT, 
	   tickets.DESCRIPTION, 
	   styles.DESCRIPTION_2, 
	   tickets.DEPT, 
	   tickets.TYP,
	   tickets.OF1, 
	   tickets.OF5,
	   tickets.PRICE,
	   buckets.LAST_COST,
	   parts.CONTACT_ID, 
	   contacts.COMPANY,
	   parts.PART_NUM,
	   parts.COST
ORDER BY [Style];