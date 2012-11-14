(ns cloxls.writer-test
    (:require
      [cloxls.writer :as w]))

;(let [wb (w/create-workbook "test.xls")
;      sheet (w/create-sheet! wb)]
;  (doto sheet
;        (w/add-cell-data! 0 0 "Company")
;        (w/add-cell-data! 0 1 "Number of Employees")
;        (w/add-cell-data! 1 0 "Dog show")
;        (w/add-cell-data! 1 1 4)
;        )
;  (w/close-workbook! wb)
;  )

(with-open [wb (w/create-workbook "test.xls")]
  (w/create-sheet! wb)
  (w/with-sheet wb 0
    (w/add-cell-data! 0 0 "Company")
    (w/add-cell-data! 0 1 "Number of Employees")
    (w/add-cell-data! 1 0 "Dog show")
    (w/add-cell-data! 1 1 10)
    )
  )

(with-open [wb (w/create-workbook "test2.xls")]
  (w/create-sheet! wb)
  (w/with-sheet wb 0
    (w/add-line-data! 0 ["Company" "Number of Employees"])
    (w/add-line-data! 1 ["Dog show" 15])))

(with-open [wb (w/create-workbook "test3.xls")]
  (w/create-sheet! wb)
  (w/with-sheet wb 0
    (w/add-2d-data! [["Company" "Number of Employees"]
                     ["Dog show" 20]
                     ["Lazy comp" 5]
                     ["Total" "=SUM(B2,B3)"]
                     ])))


