SET @tValue = 0.978; #80% T Value @ N=3
SET @N = 3; #Obeservations - 2  (5 years of data)

SELECT ID, Description1 as `DESCRIPTION`, 
QUANTITY, round(QUANTITY/UP_DAILY, 0) AS UP_DAY_REMAIN, round(QUANTITY/LOW_DAILY, 0) AS LOW_DAY_REMAIN,
AVG_ANNUAL_SALE, UP_ANNUAL, LOW_ANNUAL, UpperBoundAVG, LowBoundAVG,
SALE_17, SALE_18, SALE_19, SALE_20, SALE_21,
Slope, Intercept, R2, ROUND(StdError, 2) as StdEr
FROM (
	SELECT (p.OnHandQty + p.OnOrderQty - p.CommittedQty) AS QUANTITY, i.*
	FROM psi_bw.`psi ic parts_on_hand` as p
	RIGHT JOIN (
		SELECT *,
		UpperBound/365 AS UP_DAILY, UpperBound as UP_ANNUAL,
		LowBound/365 AS LOW_DAILY, LowBound as LOW_ANNUAL
		FROM (
			SELECT *,
			ROUND(1-(SSr/SSt), 4) AS R2,
      ROUND((6 * (Slope + (@tValue * StdError)) + Intercept), 2) As UpperBound,
			ROUND((6 * (Slope - (@tValue * StdError)) + Intercept), 2) AS LowBound,
      ROUND(AVG_ANNUAL_SALE + (1.638 * StdError), 2) As UpperBoundAVG,
			ROUND(AVG_ANNUAL_SALE - (1.638 * StdError), 2) AS LowBoundAVG
			FROM (    
				SELECT *,
        S / sqrt(SSxx) AS StdError,
				power(SALE_17 - AVG_ANNUAL_SALE, 2) + power(SALE_18 - AVG_ANNUAL_SALE, 2) + power(SALE_19 - AVG_ANNUAL_SALE, 2) + power(SALE_20 - AVG_ANNUAL_SALE, 2) + power(SALE_21 - AVG_ANNUAL_SALE, 2) AS SST,
				power(SALE_17 - ((Slope*1)+Intercept), 2) + power(SALE_18 - ((Slope*2)+Intercept), 2) + power(SALE_19 - ((Slope*3)+Intercept), 2) + power(SALE_20 - ((Slope*4)+Intercept), 2) + power(SALE_21 - ((Slope*5)+Intercept), 2) AS SSR 
				FROM (
					SELECT *,
					(Sy - Slope * Sx)  / 5 AS Intercept,
					sqrt((SSyy - ((SSxy/SSxx) * SSxy)) / (3)) AS S,
          SSyy - ((SSxy/SSxx)*SSxy)  as SSE
					FROM (
						SELECT *,
						((5 * Sxy) - (Sx * Sy)) / (5 * Sx2 - (power(Sx, 2))) as Slope,
						Sx2 - (power(Sx, 2)/5) as SSxx,
						Sy2 - (power(Sy, 2)/5) as SSyy,
						Sxy - ((Sx*Sy)/5) as SSxy
						FROM (
							SELECT *,
							ROUND((SALE_17 + SALE_18 + SALE_19 + SALE_20 + SALE_21)/5,2) as AVG_ANNUAL_SALE,
							15 as Sx,
							(SALE_17 + SALE_18 + SALE_19 + SALE_20 + SALE_21) AS Sy,
							(SALE_17 * 1 + SALE_18 * 2 + SALE_19 * 3 + SALE_20 * 4 + SALE_21 * 5) AS Sxy,
							55 as Sx2,
							(SALE_17 * SALE_17 + SALE_18 * SALE_18 + SALE_19 * SALE_19 + SALE_20 * SALE_20 + SALE_21 * SALE_21) AS Sy2
							FROM (
								SELECT p.*, IFNULL(d17.sale, 0) as SALE_17, IFNULL(d18.sale, 0) as SALE_18, IFNULL(d19.sale, 0) as SALE_19, IFNULL(d20.sale, 0) as SALE_20, IFNULL(d21.sale, 0) as SALE_21
								FROM (  # The follow reflects stock quantities divided anually, Union required for stock adjustments 
									(
										SELECT *
										FROM psi_bw.`psi ic parts`) as p
									LEFT JOIN (
										Select ID, sum(OriginalQty) as sale
										FROM (
											SELECT * FROM psi_bw.`psi ic issues`
											UNION
											SELECT * FROM psi_bw.`psi ic adjustments`) as u
										WHERE TransactionDate BETWEEN '2017-01-01' AND '2017-12-31'
										GROUP BY ID) AS d17
									ON p.ID = d17.ID
									LEFT JOIN (
										Select ID, sum(OriginalQty) as sale
										FROM (
											SELECT * FROM psi_bw.`psi ic issues`
											UNION
											SELECT * FROM psi_bw.`psi ic adjustments`) as u
										WHERE TransactionDate BETWEEN '2018-01-01' AND '2018-12-31'
										GROUP BY ID ) AS d18
									ON p.ID = d18.ID
									LEFT JOIN ( 
										Select ID, sum(OriginalQty) as sale
										FROM (
											SELECT * FROM psi_bw.`psi ic issues`
											UNION
											SELECT * FROM psi_bw.`psi ic adjustments`) as u
										WHERE TransactionDate BETWEEN '2019-01-01' AND '2019-12-31'
										GROUP BY ID) AS d19
									ON p.ID = d19.ID
									LEFT JOIN (
										Select ID, sum(OriginalQty) as sale
										FROM (
											SELECT * FROM psi_bw.`psi ic issues`
											UNION
											SELECT * FROM psi_bw.`psi ic adjustments`) as u
										WHERE TransactionDate BETWEEN '2020-01-01' AND '2020-12-31'
										GROUP BY ID) AS d20
									ON p.ID = d20.ID
									LEFT JOIN (
										Select ID, sum(OriginalQty) as sale
										FROM (
											SELECT * FROM psi_bw.`psi ic issues`
											UNION
											SELECT * FROM psi_bw.`psi ic adjustments`) as u
										WHERE TransactionDate BETWEEN '2021-01-01' AND '2021-12-31'
										GROUP BY ID) AS d21
									ON p.ID = d21.ID
								) 
							) AS q
							WHERE (SALE_17 != 0 OR SALE_18 != 0 OR SALE_19 != 0 OR SALE_20 != 0 OR SALE_21 != 0) 
							AND (ID NOT LIKE 'TST%' AND ID NOT LIKE 'LBL%' AND ID NOT LIKE 'LRS%' AND ID NOT LIKE 'BRP%' AND ID NOT LIKE 'CYL%' AND ID NOT LIKE 'ISS%')
						) AS sum
					) AS ss
				) AS StdEr
			) AS Bound
		) AS R2
	) AS i
	ON p.ID = i.ID
) AS Final
