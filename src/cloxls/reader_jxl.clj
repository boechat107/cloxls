(ns cloxls.reader
    "Functions to read a XLS files. If the file has formulas, the calculated values are
     gotten.
     References:
        http://www.vogella.com/articles/JavaExcel/article.html
        Formulas.java (demo from the library source code)"
    (:import
      [jxl Cell CellType FormulaCell Sheet Workbook]
      )
    )


(defn count-sheets
  "Gets the number of sheets in the given workbook."
  [wb]
  (.getNumberOfSheets wb))


(defn- get-row-contents
  "Gets the contents of a row and return them as a vector."
  [sheet row-idx]
  (let [row-cells (.getRow sheet row-idx)])
  )


(defn get-sheet-contents
  "Gets the contents of specific sheet of a workbook. Formulas are gotten as calculated 
   values if it is possible.
   The sheet-id must be a integer (index) or a string (name)."
  [wb sheet-id]
  (let [sheet (.getSheet wb sheet-id)]
    (reduce (fn [res row-idx]
                (reduce #()
                        nil
                        (range (.getColumns sheet))
                        )
                )
            nil
            (range (.getRows sheet))
            )
    )
  )
