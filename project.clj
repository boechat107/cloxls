(defproject org.clojars.boechat107/cloxls "0.3.0-SNAPSHOT"
  :description "Clojure library to write and read XLS files."
  :url "https://github.com/boechat107/cloxls"
  :license {:name "Eclipse Public License"
            :url "http://www.eclipse.org/legal/epl-v10.html"}
  :dependencies [[org.clojure/clojure "1.5.1"]
                 [org.apache.poi/poi "3.9"]]
  :profiles {:dev {:resource-paths ["resources"]}})
