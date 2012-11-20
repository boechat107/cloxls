(ns cloxls.writer
    "Functions to write data to a XLS file.
     References:
        http://www.vogella.com/articles/JavaExcel/article.html
        http://www.andykhan.com/jexcelapi/tutorial.html#writing"
    (:import
      [java.util Locale]
      [java.io File IOException]
      [jxl CellView Workbook WorkbookSettings]
      [jxl.write Formula Label WritableCellFormat WritableFont WritableSheet
                 WritableWorkbook WriteException]
      [jxl.write.biff RowsExceededException]
      )
    )

(defn create-workbook
  "Creates and returns the high level abstraction of the spreadsheet, a WritableWorkbook
   object. The default Locale is EN." 
  [filename]
  (let [file (File. filename)
        wb-set (WorkbookSettings.)]
    (try
      (Workbook/createWorkbook file)
      (catch IOException e
             (str "Problem to create a xls file: " (.getMessage e))))))


;; ============== Sheet contents ==============================


(def ^{:dynamic true
       :doc "This variable is bound to a sheet when the WITH-SHEET macro is used."}
       *sheet* nil)


(defn create-sheet!
  "Creates a sheet with the given name or with a default name. Side effects only." 
  ([wb] 
   (let [n (inc (.getNumberOfSheets wb))]
     (create-sheet! wb (str "Sheet " n) n)))
  ([wb sheet-name idx]
   (let [sheet (.createSheet wb sheet-name idx)
         settings (.getSettings sheet)]
     sheet)))


;(defn add-label!
;  "Add a text label to a specific cell."
;  [sheet row col text]
;  (.addCell sheet (Label. row col text)))
;
;(defn add-number!
;  "Add a number to a specific cell."
;  [sheet row col num]
;  (.addCell sheet (jxl.write.Number. row col num)))


(defn add-cell-data!
  "Adds data to a specific cell. If the data is a string, a label is created. Otherwise, a
   number is created. A formula is created using a string where the first character is =."
  ([row col data] (add-cell-data! *sheet* row col data))
  ([sheet row col data]
   (let [dtype (cond
                 (number? data) (jxl.write.Number. col row data)
                 (and (string? data) (= \= (get data 0))) (Formula. col row (subs data 1))
                 :else (Label. col row data))]
     (.addCell sheet dtype))))


(defn- coll-idx-data
  "Returns a new collection composed of tuples (idx, element), where idx is the index of
   the element in the original collection."
  [coll]
  (partition 2 (interleave (range (count coll)) coll)))


(defn add-line-data!
  "Adds data from a collection to a specific sheet's line, the row number."
  ([row coll] (add-line-data! *sheet* row coll))
  ([sheet row coll]
   (doseq [[col data] (coll-idx-data coll)]
     (add-cell-data! sheet row col data))))


(defn add-2d-data!
  "Adds data to a sheet as a matrix, starting from line 0 and column 0."
  ([mat] (add-2d-data! *sheet* mat))
  ([sheet mat]
   (doseq [[row line] (coll-idx-data mat)]
     (add-line-data! sheet row line))))


(defmacro with-sheet
  "Binds the variable *sheet* to a sheet with the given sheet-id of the workbook wb,
   exposes it to the body and write the modifications to file."
  [wb sheet-id & body]
  `(let [wb# ~wb]
     (binding [*sheet* (.getSheet wb# ~sheet-id)]
       ~@body
       (.write wb#))))


(defn close-workbook!
  [wb]
  (doto wb (.write) (.close)))
