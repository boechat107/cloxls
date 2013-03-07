(ns cloxls.core
    "References:
        http://poi.apache.org/spreadsheet/quick-guide.html"
  (:import 
   [java.io IOException FileOutputStream FileInputStream]
   [org.apache.poi.hssf.usermodel HSSFWorkbook HSSFCell HSSFSheet HSSFRow]
   [org.apache.poi.ss.util CellRangeAddress]
   [org.apache.poi.poifs.filesystem POIFSFileSystem]))


(defonce ^{:dynamic true
           :doc "This variable is bound to the created workbook when the macro
                  with-new-wb is used." }
  *wb* nil)


(defonce ^{:dynamic true
           :doc "This variable is bound to a sheet when the WITH-SHEET macro is used."}
  *sheet* nil)

;; ============== Workbook manipulation ==============================

(defmacro with-new-wb
  "Binds the variable *new-wb* to a new workbook whose name is given."
  [filename & body]
  `(let [file# (FileOutputStream. ~filename)]
     (binding [*wb* (HSSFWorkbook.)]
       ~@body
       (.write *wb* file#)
       (.close file#))))


(defmacro with-wb
  "Binds the variable  *new-wb* to a existing workbook whose name is given.
   Use this macro just for reading."
  [filename & body]
  `(let [file# (FileInputStream. ~filename)]
     (binding [*wb* (HSSFWorkbook. (POIFSFileSystem. file#))]
       (let [res# (do ~@body)] 
         (.close file#)
         res#))))

;; ============== Sheet manipulation ==============================

(defn create-sheet!
  "Creates a sheet with the given name or with a default name. The integer index of the
   created sheet is returned."
  ([] (create-sheet! *wb*))
  ([wb] 
   (let [n (inc (.getNumberOfSheets wb))]
     (create-sheet! (str "Sheet " n) wb)))
  ([sheet-name wb]
   (.createSheet wb sheet-name)
   (.getSheetIndex wb sheet-name)))


(defmacro with-sheet
  "Binds the variable *sheet* to a sheet with the given sheet-id of the workbook *wb*,
   exposes it to the body and write the modifications to file. "
  [sheet-id & body]
  {:pre [(integer? sheet-id)]}
  `(let [sheet-id# ~sheet-id]
     (binding [*sheet* (if (number? sheet-id#)
                         (.getSheetAt *wb* sheet-id#)
                         (.getSheet *wb* sheet-id#))]
      (let [res# (do ~@body)]
        ;; Auto resize the columns' width.
;        (doseq [c# (range (get-num-cols))]
;          (.autoSizeColumn *sheet* c#))
        res#))))

;; ============== Adding cell contents ==============================

(defn create-cell!
  "Creates a cell, adds data to it and sets it on the sheet/workbook.
   The data could be a number or a string. If the string starts with =, a formula is
   created."
  [row-obj c-id data] 
   {:pre [(instance? HSSFRow row-obj) (integer? c-id)]}
   (let [cell (.createCell row-obj c-id)]
     (cond
       ;; Formula cell, string whose first character is =.
       (and (string? data)
            (= \= (get data 0))) (.setCellFormula cell (subs data 1))
       ;; Number cell.
       (number? data) (.setCellValue cell (double data))
       ;; Label cell, a string.
       :else (.setCellValue cell (str data)))))


(defn- coll-idx-data
  "Returns a range of numbers corresponding to the index of each element of coll. A offset
   can be used as a starting number."
  [coll start]
  (range start (+ start (count coll))))


(defn create-row-data!
  "Adds a entire or partial row of data from a data collection to a sheet."
  ([r-id data] (create-row-data! *sheet* r-id 0 data))
  ([r-id c-start data] (create-row-data! *sheet* r-id c-start data))
  ([sheet r-id c-start data]
   {:pre [(instance? HSSFSheet sheet) (integer? r-id) (integer? c-start) (coll? data)]} 
     (let [row-obj (or (.getRow sheet r-id) (.createRow sheet r-id))]
       (doseq [[c-id d] (map vector (coll-idx-data data c-start) data)]
         (create-cell! row-obj c-id d)))))


(defn create-col-data!
  "Creates a column of data at the column c-id, starting at the row r-start."
  ([c-id data] (create-col-data! *sheet* c-id 0 data))
  ([c-id r-start data] (create-col-data! *sheet* c-id r-start data))
  ([sheet c-id r-start data] 
   {:pre [(instance? HSSFSheet sheet) (integer? c-id) (integer? r-start) (coll? data)]} 
   (let [rows-id (coll-idx-data data r-start)
         get-row (fn [r] (or (.getRow sheet r) (.createRow sheet r)))]
     (doseq [[r-obj d] (map vector (map get-row rows-id) data)]
       (create-cell! r-obj c-id d)))))


(defn create-2d-data!
  "Adds the data to a sheet as a matrix, starting from line r-start and column c-start."
  ([data] (create-2d-data! *sheet* 0 0 data))
  ([r-start c-start data] (create-2d-data! *sheet* r-start c-start data))
  ([sheet r-start c-start data]
   {:pre [(instance? HSSFSheet sheet) (integer? r-start) (integer? c-start) (coll? data)]}
   (doseq [[r-id ld] (map vector (coll-idx-data data r-start) data)]
     (create-row-data! sheet r-id c-start ld))))


;; ============== Reading cell contents ==============================

(defn get-cell-content
  "Gets the content of the cell considering its type. If formula-eval? is true, the cell 
   is evaluated and formulas results are returned instead."
  ;; TODO: support date format.
  ([cell formula-eval?] (get-cell-content *wb* cell formula-eval?))
  ([wb cell formula-eval?]
   (let [eval-type (when formula-eval?
                     (-> (.getCreationHelper wb)
                         (.createFormulaEvaluator)
                         (.evaluateFormulaCell cell)))
         cell-type (if (or (nil? eval-type) (= -1 eval-type))
                     (.getCellType cell)
                     eval-type)]
     (condp = cell-type
       HSSFCell/CELL_TYPE_STRING (-> (.getRichStringCellValue cell)
                                     (.getString))
       HSSFCell/CELL_TYPE_NUMERIC (.getNumericCellValue cell)
       HSSFCell/CELL_TYPE_BOOLEAN (.getBooleanCellValue cell)
       HSSFCell/CELL_TYPE_FORMULA (.getCellFormula cell)))))


(defn sheet->matrix
  "Gets the contents of specific sheet of a workbook. Formulas are returned as calculated 
   values if it is possible.
   The sheet-id must be a integer (index) or a string (name).
   If formula-eval? is true, the cells are evaluated and formulas results are returned
   (default false)."
  ([sheet-id] (sheet->matrix *wb* sheet-id nil))
  ([sheet-id formula-eval?] (sheet->matrix *wb* sheet-id formula-eval?))
  ([wb sheet-id formula-eval?]
   {:pre [(instance? HSSFWorkbook wb) (or (string? sheet-id) (integer? sheet-id))]}
     (let [sheet (if (number? sheet-id)
                   (.getSheetAt wb sheet-id)
                   (.getSheet wb sheet-id))]
       (vec (map (fn [row-obj]
                   (vec (map #(let [cell (.getCell row-obj %)]
                                (when cell
                                  (get-cell-content wb cell formula-eval?)))
                             (range (.getLastCellNum row-obj)))))
                 (iterator-seq (.rowIterator sheet)))))))


(defn get-num-cols
  "Iterates over all rows of a sheet and returns the maximum number of columns of all 
   rows."
  ([] (get-num-cols *sheet*))
  ([sheet]
   {:pre [(instance? HSSFSheet sheet)]}
   (->> (.rowIterator sheet)
        (iterator-seq)
        (map #(.getLastCellNum %))
        (reduce max))))

;; ============== Format functions ==============================

(defn autosize-columns!
  "Resize the columns width to fit its contents. If the number of columns is given, this 
   function works much faster for big tables."
  ([] (autosize-columns! *sheet* (get-num-cols)))
  ([n-cols] (autosize-columns! *sheet* n-cols))
  ([sheet n-cols]
   (doseq [c (range n-cols)]
     (.autoSizeColumn sheet c))))

(defn- get-color-idx
  "Gets the integer index of a color.
   Possible colors:
   http://poi.apache.org/apidocs/org/apache/poi/hssf/util/HSSFColor.html"
  [color]
  {:pre [(keyword? color)]}
  (let [prefix "org.apache.poi.hssf.util.HSSFColor$"
        c-str (.toUpperCase (name color))]
    (eval (symbol (str prefix c-str "/index")))))

(defn- formatting-rule
  "Returns a ConditionalFormattingRule obeject from a SheetConditionalFormatting
   object (sheet-cf), a rule (string) and a map of objects to be formated."
  [sheet-cf rule-map]
  (let [rule (:rule rule-map)
        font-conf (:font rule-map)
        cf-rule (.createConditionalFormattingRule sheet-cf rule)]
    (when font-conf
      (doto (.createFontFormatting cf-rule)
            (.setFontStyle (or (:italic font-conf) false) 
                           (or (:bold font-conf) false))
            (.setFontColorIndex (get-color-idx (:color font-conf)))))
    cf-rule))

(defn conditional-formatting!
  "Formats a region (a simple string) or regions (a seq of strings) of cells
   following a rule (string). objs is a map of cell components that can be modified.  
   EXAMPLES: 
    (conditional-formatting! \"B1:B10\"
                             {:rule \"A1>10\", :font {:color :blue}})
    (conditional-formatting! [\"A4:B4\" \"A1:B1\"]
                             [{:rule \"$B$2>10\", :font {:color :green}}
                              {:rule \"$B$2<=10\", :font {:color :blue}}])
    :font options
          :color  :blue, :green, :black ... (see http://poi.apache.org/apidocs/org/apache/poi/hssf/util/HSSFColor.html)
          :bold   true, false
          :italic true, false"
  ([regions f-map] (conditional-formatting! *sheet* regions f-map))
  ([sheet regions f-map]
   {:pre [(instance? HSSFSheet sheet) (or (coll? regions) (string? regions)) 
          (or (coll? f-map) (map? f-map))]} 
   (let [sheet-cf (.getSheetConditionalFormatting sheet)
         reg-array (->> (if (coll? regions) regions [regions])
                        (map #(CellRangeAddress/valueOf %))
                        (into-array))]
     (->> (map #(formatting-rule sheet-cf %) (if (coll? f-map) f-map [f-map]))
          (into-array)
          (.addConditionalFormatting sheet-cf reg-array))
     nil)))
