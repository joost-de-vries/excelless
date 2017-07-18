import java.io.InputStream
import java.time.{Instant, ZoneId}
import java.util.TimeZone

import org.apache.poi.ss.usermodel._
import shapeless._

import scala.collection.immutable.{:: => Cons}
import scala.util.{Failure, Success, Try}

// The class to serialize or deserialize
case class Person(name: String,
                  middleName: Option[String],
                  surname: String,
                  birthdate: Instant,
                  age: Try[Int],
                  id: Option[Int])

object ExcelExample extends App {
  import ExcelReader._
  implicit val timezone = TimeZone.getTimeZone("UTC")

  val persons = fromResource[Person]("Workbook1.xlsx")(wb => wb.getSheetAt(0))
  persons.foreach(println)

}

object ExcelReader {
  def fromResource[A](resourceName: String)(getSheet: Workbook => Sheet)(implicit conv: ExcelConverter[List[Cell], A]) =
    Try(getClass.getResourceAsStream(resourceName))
      .flatMap(fromInputStream[A](getSheet))

  def fromInputStream[A](getSheet: Workbook => Sheet)(is: InputStream)(
      implicit conv: ExcelConverter[List[Cell], A]): Try[List[A]] = {
    Try {
      val wb = WorkbookFactory.create(is)

      val sheet = getSheet(wb)

      val tries = parseSheet(sheet).map(r => parseRow(r))
      tries.foreach(t => println(s"try ${t}"))
      tries.collect {
        case Success(a) => a
      }
    }
  }

  def parseRow[A](row: Row)(implicit conv: ExcelConverter[List[Cell], A]): Try[A] = {
    import collection.JavaConverters._

    ExcelConverter[List[Cell], A].from(row.cellIterator().asScala.toList)
  }

  def parseSheet(sheet: Sheet): List[Row] = {
    import collection.JavaConverters._
    sheet.rowIterator().asScala.takeWhile(_.getFirstCellNum != -1).toList
  }
}

// Implementation

class ExcelException(s: String) extends RuntimeException(s)

// based on https://github.com/milessabin/shapeless/blob/master/examples/src/main/scala/shapeless/examples/csv.scala
trait ExcelConverter[S, T] {
  def from(s: S): Try[T]

  //  def to(t: T): String
}

object ExcelConverter {
  def apply[S, T](implicit st: Lazy[ExcelConverter[S, T]]): ExcelConverter[S, T] = st.value

  def fail(s: String): Failure[Nothing] = Failure(new ExcelException(s))

  // Primitives
  implicit def dateConverter(implicit timeZone: TimeZone): ExcelConverter[Cell, Instant] =
    new ExcelConverter[Cell, Instant] {
      def from(cell: Cell): Try[Instant] = cell.getCellTypeEnum match {

        case CellType.NUMERIC if DateUtil.isCellDateFormatted(cell) =>
          Success(DateUtil.getJavaDate(cell.getNumericCellValue, timeZone).toInstant);
        case cellType =>
          Failure(
            new IllegalArgumentException(
              s"${cell.getAddress} expected dateformatted numeric received $cellType ${cell.toString}"))
      }

      //    def to(i: Boolean): String = i.toString
    }

  implicit def stringExcelConverter: ExcelConverter[Cell, String] = new ExcelConverter[Cell, String] {
    def from(cell: Cell): Try[String] = cell.getCellTypeEnum match {
      case CellType.STRING => Success(cell.getStringCellValue)
      case cellType =>
        Failure(
          new IllegalArgumentException(
            s"${cell.getAddress} expected string received $cellType ${cell.toString} is datetime ${DateUtil
              .isCellDateFormatted(cell)}"))
    }

    //    def to(s: String): String = s
  }

  implicit def intConverter: ExcelConverter[Cell, Int] = new ExcelConverter[Cell, Int] {
    def from(cell: Cell): Try[Int] = cell.getCellTypeEnum match {

      case CellType.NUMERIC => Success(cell.getNumericCellValue.toInt)
      case cellType =>
        Failure(new IllegalArgumentException(s"${cell.getAddress} expected int received $cellType ${cell.toString}"))
    }

    //    def to(i: Int): String = i.toString
  }

  implicit def booleanConverter: ExcelConverter[Cell, Boolean] = new ExcelConverter[Cell, Boolean] {
    def from(cell: Cell): Try[Boolean] = cell.getCellTypeEnum match {

      case CellType.BOOLEAN => Success(cell.getBooleanCellValue)
      case cellType =>
        Failure(
          new IllegalArgumentException(s"${cell.getAddress} expected boolean received $cellType ${cell.toString}"))
    }

    //    def to(i: Boolean): String = i.toString
  }

  def listCsvLinesConverter[A](l: List[List[Cell]])(implicit ec: ExcelConverter[List[Cell], A]): Try[List[A]] =
    l match {
      case Nil => Success(Nil)
      case Cons(s, ss) =>
        for {
          x  <- ec.from(s)
          xs <- listCsvLinesConverter(ss)(ec)
        } yield Cons(x, xs)
    }

  implicit def listCsvConverter[A](implicit ec: ExcelConverter[List[Cell], A]): ExcelConverter[Sheet, List[A]] =
    new ExcelConverter[Sheet, List[A]] {

      import ExcelReader.parseSheet

      import collection.JavaConverters._

      def from(s: Sheet): Try[List[A]] = listCsvLinesConverter(parseSheet(s).map(_.cellIterator().asScala.toList))(ec)

      // def to(l: List[A]): String = l.map(ec.to).mkString("\n")
    }

  // HList
  implicit def deriveHNilForList: ExcelConverter[List[Cell], HNil] =
    new ExcelConverter[List[Cell], HNil] {
      def from(cells: List[Cell]): Try[HNil] = cells match {
        case Nil => Success(HNil)
        case _   => fail("Cannot convert '" ++ cells.toString ++ "' to HNil")
      }

      //      def to(n: HNil) = ""
    }

  implicit def deriveHCons[V, T <: HList](
      implicit scv: Lazy[ExcelConverter[Cell, V]],
      sct: Lazy[ExcelConverter[List[Cell], T]]): ExcelConverter[List[Cell], V :: T] =
    new ExcelConverter[List[Cell], V :: T] {

      def from(r: List[Cell]): Try[V :: T] = r match {
        case Cons(before, after) =>
          for {
            front <- scv.value.from(before)
            back  <- sct.value.from(after)
          } yield front :: back

        case _ => fail("Cannot convert '" ++ r.toString ++ "' to HList")
      }

      //      def to(ft: V :: T): String = {
      //        scv.value.to(ft.head) ++ "," ++ sct.value.to(ft.tail)
      //      }
    }

  implicit def deriveHConsOption[V, T <: HList](
      implicit scv: Lazy[ExcelConverter[Cell, V]],
      sct: Lazy[ExcelConverter[List[Cell], T]]): ExcelConverter[List[Cell], Option[V] :: T] =
    new ExcelConverter[List[Cell], Option[V] :: T] {
      override def from(r: List[Cell]): Try[::[Option[V], T]] = {
        r match {
          case Cons(before, after) =>
            (for {
              front <- scv.value.from(before)
              back  <- sct.value.from(after)
            } yield Some(front) :: back).orElse {
              sct.value.from(r).map(None :: _) // r or after
            }

          case List() => // necessary?
            for {
              back <- sct.value.from(Nil)
            } yield None :: back

          case _ => fail("Cannot convert '" ++ r.toString ++ "' with Option to HList")
        }
      }
    }

  implicit def deriveHConsTry[V, T <: HList](
      implicit scv: Lazy[ExcelConverter[Cell, V]],
      sct: Lazy[ExcelConverter[List[Cell], T]]): ExcelConverter[List[Cell], Try[V] :: T] =
    new ExcelConverter[List[Cell], Try[V] :: T] {
      override def from(r: List[Cell]): Try[::[Try[V], T]] = {
        r match {
          case Cons(before, after) =>
            (for {
              front <- scv.value.from(before)
              back  <- sct.value.from(after)
            } yield Success(front) :: back).orElse {
              sct.value.from(after).map(scv.value.from(before) :: _)
            }

          case _ => fail("Cannot convert '" ++ r.toString ++ "' with Option to HList")
        }
      }
    }

  // generics
  implicit def deriveClass[A, R](implicit gen: Generic.Aux[A, R],
                                 conv: ExcelConverter[List[Cell], R]): ExcelConverter[List[Cell], A] =
    new ExcelConverter[List[Cell], A] {

      def from(row: List[Cell]): Try[A] = conv.from(row).map(gen.from)

      //    def to(a: A): String = conv.to(gen.to(a))
    }
}
