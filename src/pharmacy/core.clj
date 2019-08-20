(ns pharmacy.core
  (:require
   [clojure.string :as string]
   [dk.ative.docjure.spreadsheet :as spreadsheet]
   [clojure.java.io :as io])
  (:gen-class))

(def ^:private columns-map
  {:A :odr_code
   :B :t_01
   :C :t_02
   :D :t_03
   :E :t_04
   :F :t_05
   :G :t_06
   :H :t_07
   :I :t_08
   :J :t_09
   :K :t_10
   :L :t_11
   :M :t_12
   :N :t_13
   :O :t_14
   :P :t_15
   :Q :t_16
   :R :t_17
   :S :t_18
   :T :t_19
   :U :t_20
   :V :t_21
   :W :t_22
   :X :t_23
   :Y :t_24
   :Z :t_25
   :AA :t_26
   :AB :t_27
   :AC :t_28
   :AD :t_29})

(def template
  (str "UPDATE [ESMC_TextData] "
       "SET [textinfo_01] = %s,"
       " [textinfo_02] = %s,"
       " [textinfo_03] = %s,"
       " [textinfo_04] = %s,"
       " [textinfo_05] = %s,"
       " [textinfo_06] = %s,"
       " [textinfo_07] = %s,"
       " [textinfo_08] = %s,"
       " [textinfo_09] = %s,"
       " [textinfo_10] = %s,"
       " [textinfo_11] = %s,"
       " [textinfo_12] = %s,"
       " [textinfo_13] = %s,"
       " [textinfo_14] = %s,"
       " [textinfo_15] = %s,"
       " [textinfo_16] = %s,"
       " [textinfo_17] = %s,"
       " [textinfo_18] = %s,"
       " [textinfo_19] = %s,"
       " [textinfo_20] = %s,"
       " [textinfo_21] = %s,"
       " [textinfo_22] = %s,"
       " [textinfo_23] = %s,"
       " [textinfo_24] = %s,"
       " [textinfo_25] = %s,"
       " [textinfo_26] = %s,"
       " [textinfo_27] = %s,"
       " [textinfo_28] = %s,"
       " [textinfo_29] = %s"
       " WHERE [odr_code] = %s"))

(defn escape
  [d-str]
  (if (= \' (first d-str))
    (string/escape (apply str (rest d-str)) {\' "''"})
    (string/escape d-str {\' "''"})))

(defn c-nil
  [d]
  (if (nil? d)
    "NULL"
    (if (string? d)
      (format "'%s'" (escape d))
      (format "'%s'" d))))

(defn data->str
  [data]
  (let [d-l [(:t_01 data)
             (:t_02 data)
             (:t_03 data)
             (:t_04 data)
             (:t_05 data)
             (:t_06 data)
             (:t_07 data)
             (:t_08 data)
             (:t_09 data)
             (:t_10 data)
             (:t_11 data)
             (:t_12 data)
             (:t_13 data)
             (:t_14 data)
             (:t_15 data)
             (:t_16 data)
             (:t_17 data)
             (:t_18 data)
             (:t_19 data)
             (:t_20 data)
             (:t_21 data)
             (:t_22 data)
             (:t_23 data)
             (:t_24 data)
             (:t_25 data)
             (:t_26 data)
             (:t_27 data)
             (:t_28 data)
             (:t_29 data)
             (:odr_code data)]
        d-l-no-nil (mapv c-nil d-l)]
    (apply format template d-l-no-nil)))

(defn output [cont]
  (spit "sql.txt" (str cont "\n") :append true))

(defn process [raw]
  (->> (map data->str raw)
       (map output)))

(defn get-raw-from-excel-fn*
  "assemble columns-map and get-raw-from-excel fn
   The * means `get` title and data"
  [columns-map]
  (fn get-raw-from-excel
    [addr filename]
    (try
      (with-open [stream (io/input-stream (str addr filename))]
        (let [title+orders (->> (spreadsheet/load-workbook stream)
                                (spreadsheet/select-sheet "Sheet0")
                                (spreadsheet/select-columns columns-map))]
          title+orders))
      (catch Exception e
        (let [desc "get-raw-from-excel error"]
          (throw (ex-info desc (Throwable->map e))))))))

(defn get-raw-from-excel-fn
  "assemble columns-map and get-raw-from-excel fn"
  [columns-map]
  (comp rest (get-raw-from-excel-fn* columns-map)))

(def ^:private get-raw-from-excel
  (get-raw-from-excel-fn columns-map))

(comment
  (def raw (get-raw-from-excel "http://10.20.30.40:5001/" "data.xlsx")))

(defn -main
  "I don't do a whole lot ... yet."
  [& args]
  (println "Hello, World!"))
