name := "excel"

version := "1.0"

scalaVersion := "2.11.8"

libraryDependencies ++= Seq(
  "org.apache.poi" % "poi" % "3.15",
  "org.apache.poi" % "poi-ooxml" % "3.15",
  "com.chuusai" %% "shapeless" % "2.3.2"
)