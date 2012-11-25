(ns cloxls.reader-jxl
    "Functions to read a XLS files. If the file has formulas, the calculated values are
     gotten.
     References:
        http://www.vogella.com/articles/JavaExcel/article.html
        Formulas.java (demo from the library source code)"
    (:import
      [jxl Cell CellType NumberCell NumberFormulaCell FormulaCell Sheet Workbook]
      [java.io File]
      )
    )


(defn count-sheets
  "Gets the number of sheets in the given workbook."
  [wb]
  (.getNumberOfSheets wb))


(defn- cell-formula?
  [cell]
  (let [ct (.getType cell)]
    (or (= CellType/NUMBER_FORMULA ct)
        (= CellType/STRING_FORMULA ct)
        (= CellType/BOOLEAN_FORMULA ct)
        (= CellType/DATE_FORMULA ct)
        (= CellType/FORMULA_ERROR ct))))


(defn- get-cell-content
  [cell]
  (cond
    (= CellType/NUMBER (.getType cell)) (-> (.getContents cell)
                                            (read-string))
    (cell-formula? cell) (-> ^NumberCell cell
                             (.getValue))
    :else (.getContents cell)))


(defn get-sheet-contents
  "Gets the contents of specific sheet of a workbook. Formulas are gotten as calculated 
   values if it is possible.
   The sheet-id must be a integer (index) or a string (name)."
  [wb sheet-id]
  (let [sheet (.getSheet wb sheet-id)]
    (vec (map (fn [row-idx]
                  (vec (map #(-> (.getCell sheet % row-idx)
                                 (get-cell-content))
                            (range (.getColumns sheet)))))
              (range (.getRows sheet))
              ))
    )
  )



;(try (let [wb (Workbook/getWorkbook (File. "test2.xls"))]
;       (get-sheet-contents wb 0))
;     (catch Exception e
;            (println (.getMessage e))))
