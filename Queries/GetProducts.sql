SELECT tickets.BRAND AS [Brand], 
	   tickets.STYLE AS [Style],
	   tickets.LOOKUP AS [Barcode Lookup], 
	   tickets.STORE_ID,
	   tickets.PRICE AS [Old Retail], 
	   CAST(styles.MARGIN_PERCENT AS DECIMAL(18, 2)) AS [Old Margin], 
	   tickets.DESCRIPTION AS [Description 1], 
	   styles.DESCRIPTION_2 AS [Description 2], 
	   tickets.DEPT AS [Department], 
	   tickets.TYP AS [Type],
	   tickets.OF1 AS [SEASON], 
	   tickets.OF5 AS [PROMO],
	   buckets.QOH AS [Available],
	   tickets.PRICE AS [Last Price],
	   buckets.LAST_COST AS [Last Cost],
	   parts.CONTACT_ID, 
	   (CASE contacts.COMPANY
			WHEN 'ZEIGLER`S DISTRIBUTIOR INC' THEN 'Zeig'
			WHEN 'PHILLIP`S PET SUPPLY' THEN 'Phil'
			WHEN 'NATURAL ANIMAL NUTRITION' THEN 'NAN'
			WHEN 'BRADLEY CALDWELL' THEN 'Brad'
			WHEN 'PET FOOD EXPERTS' THEN 'PFX'
			WHEN 'TICKNER`S INC.' THEN 'Tick'
			ELSE contacts.COMPANY
		END) AS VENDOR,
		parts.COST AS [Part Cost]
FROM VW_TICKETS tickets
INNER JOIN TB_STYLES styles
ON styles.STYLE_ID = tickets.STYLE_ID
INNER JOIN TB_SKU_BUCKETS buckets
ON buckets.SKU_ID = tickets.SKU_ID
AND buckets.SKU_BUCKET_ID = tickets.SKU_BUCKET_ID
AND buckets.STORE_ID = tickets.STORE_ID
INNER JOIN TB_PARTS parts
ON parts.STYLE_ID = tickets.STYLE_ID
AND parts.CONTACT_ID <> 195
INNER JOIN TB_CONTACTS contacts
ON contacts.CONTACT_ID = parts.CONTACT_ID
WHERE tickets.BRAND = $(Brand)
AND tickets.OF1 <> 'DISCO'
AND styles.STATUS_FINISH <> 'Y';