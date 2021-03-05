inv_history_wkly = """
WITH CALENDAR AS (
    SELECT TRUNC(NEXT_DAY(SYSDATE - 7, 'SUN')) - (7 * (ROWNUM - 2)) DATETIME
    FROM DUAL
    CONNECT BY LEVEL <= 53
)

SELECT CN.WEEK,
       CN.ITEMNUM,
       CN.LOCATION,
       USE.QUANTITY                                            AS ISSUE,
       REC.QUANTITY                                            AS TRANSFER,
       COALESCE(-USE.QUANTITY, 0) + COALESCE(-REC.QUANTITY, 0) AS NET_USAGE, RCT.QUANTITY AS RECEIVED_QTY
FROM (SELECT TRUNC(DATETIME, 'IW') AS WEEK,
             ITEMNUM,
             LOCATION
      FROM MXRADS.INVENTORY,
           CALENDAR
      WHERE LOCATION = :loc_id and ROWNUM <= 10) CN

         LEFT JOIN (
    SELECT TRUNC(ACTUALDATE, 'IW')                                   AS WEEK,
           ITEMNUM,
           STORELOC,
           SITEID,
           CASE WHEN SUM(QUANTITY) > 0 THEN 0 ELSE SUM(QUANTITY) END AS QUANTITY
    FROM MSCRADS.MATUSETRANS
    WHERE LINETYPE = 'ITEM' AND LOCATION = :loc_id
      AND CONSIGNMENT = 0
      AND SITEID IN ('DIS', 'TRN')
      AND TRUNC(ACTUALDATE, 'IW') >= TRUNC(SYSDATE - 366, 'IW')
    GROUP BY TRUNC(ACTUALDATE, 'IW'),
             ITEMNUM,
             STORELOC,
             SITEID
) USE
                   ON USE.ITEMNUM = CN.ITEMNUM
                       AND USE.STORELOC = CN.LOCATION
                       AND USE.WEEK = CN.WEEK
         LEFT JOIN (
    SELECT TRUNC(ACTUALDATE, 'IW') AS WEEK,
           ITEMNUM,
           FROMSTORELOC,
           FROMSITEID,
           SUM(-QUANTITY)          AS QUANTITY
    FROM MSCRADS.MATRECTRANS
    WHERE ISSUETYPE IN (
                        'SHIPRECEIPT',
                        'SHIPRETURN',
                        'VOIDSHIPRECEIPT'
        )
      AND CONSIGNMENT = 0 AND FROMSTORELOC = :loc_id
      AND FROMSITEID IN ('DIS', 'TRN')
      AND STATUS = 'COMP'
      AND LINETYPE = 'ITEM'
      AND TRUNC(ACTUALDATE, 'IW') >= TRUNC(SYSDATE - 366, 'IW')
    GROUP BY TRUNC(ACTUALDATE, 'IW'),
             ITEMNUM,
             FROMSTORELOC,
             FROMSITEID
) REC
                   ON REC.ITEMNUM = CN.ITEMNUM
                       AND REC.FROMSTORELOC = CN.LOCATION
                       AND REC.WEEK = CN.WEEK

 LEFT JOIN (
    SELECT TRUNC(ACTUALDATE, 'IW') AS WEEK,
           ITEMNUM,
           TOSTORELOC,
           SITEID,
           SUM(QUANTITY)          AS QUANTITY
    FROM MSCRADS.MATRECTRANS
    WHERE ISSUETYPE IN (
                        'RECEIPT',
                        'RETURN',
                        'VOIDRECEIPT'
        )
      AND CONSIGNMENT = 0 AND TOSTORELOC = {location}
      AND STATUS = 'COMP'
      AND LINETYPE = 'ITEM'
      AND TRUNC(ACTUALDATE, 'IW') >= TRUNC(SYSDATE - 366, 'IW')
    GROUP BY TRUNC(ACTUALDATE, 'IW'),
             ITEMNUM,
             TOSTORELOC,
             SITEID
) RCT
                   ON RCT.ITEMNUM = CN.ITEMNUM
                       AND RCT.TOSTORELOC = CN.LOCATION
                       AND RCT.WEEK = CN.WEEK
"""
