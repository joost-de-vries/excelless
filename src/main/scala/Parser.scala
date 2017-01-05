import java.io.InputStream
import java.time.{Instant, ZoneId}
import java.util.TimeZone

import org.apache.poi.ss.usermodel._
import shapeless._

import scala.collection.immutable.{:: => Cons}
import scala.util.{Failure, Success, Try}

// The class to serialize or deserialize
case class Person(name: String, surname: String, birthdate:Instant, age: Int, id: Option[Int])

object ExcelExample extends App {
  import ExcelReader._
  implicit val timezone = TimeZone.getTimeZone("UTC")

  val persons =fromResource[Person]("Workbook1.xlsx")(wb => wb.getSheetAt(0))
  persons.foreach(println)

}

object ExcelReader {
  def fromResource[A](resourceName: String)(getSheet:Workbook=> Sheet)
                     (implicit conv: ExcelConverter[List[Cell], A]) =
    Try(getClass.getResourceAsStream(resourceName))
      .flatMap(fromInputStream[A](getSheet))

  def fromInputStream[A](getSheet:Workbook=> Sheet)(is: InputStream)
                        (implicit conv: ExcelConverter[List[Cell], A]): Try[List[A]] = {
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

/** Exception to throw if something goes wrong during CSV parsing */
class ExcelException(s: String) extends RuntimeException(s)

/** Trait for types that can be serialized to/deserialized from CSV */
trait ExcelConverter[S, T] {
  def from(s: S): Try[T]

  //  def to(t: T): String
}

/** Instances of the CSVConverter trait */
object ExcelConverter {
  def apply[S, T](implicit st: Lazy[ExcelConverter[S, T]]): ExcelConverter[S, T] = st.value

  def fail(s: String) = Failure(new ExcelException(s))


  // Primitives

  implicit def stringExcelConverter: ExcelConverter[Cell, String] = new ExcelConverter[Cell, String] {
    def from(cell: Cell): Try[String] = cell.getCellTypeEnum match {
      case CellType.STRING => Success(cell.getStringCellValue)
      case _ => Failure(new IllegalArgumentException("expected string"))
    }

    //    def to(s: String): String = s
  }

  implicit def intCsvConverter: ExcelConverter[Cell, Int] = new ExcelConverter[Cell, Int] {
    def from(cell: Cell): Try[Int] = cell.getCellTypeEnum match {

      case CellType.NUMERIC => Success(cell.getNumericCellValue.toInt)
      case _ => Failure(new IllegalArgumentException("expected int"))
    }

    //    def to(i: Int): String = i.toString
  }

  implicit def booleanCsvConverter: ExcelConverter[Cell, Boolean] = new ExcelConverter[Cell, Boolean] {
    def from(cell: Cell): Try[Boolean] = cell.getCellTypeEnum match {

      case CellType.BOOLEAN => Success(cell.getBooleanCellValue)
      case _ => Failure(new IllegalArgumentException("expected boolean"))
    }

    //    def to(i: Boolean): String = i.toString
  }

  implicit def dateCsvConverter(implicit timeZone: TimeZone): ExcelConverter[Cell, Instant] = new ExcelConverter[Cell, Instant] {
    def from(cell: Cell): Try[Instant] = cell.getCellTypeEnum match {

      case CellType.NUMERIC if DateUtil.isCellDateFormatted(cell) => Success(DateUtil.getJavaDate(cell.getNumericCellValue, timeZone).toInstant);
      case _ => Failure(new IllegalArgumentException("expected boolean"))
    }

    //    def to(i: Boolean): String = i.toString
  }


  // HList
  implicit def deriveHNilForList: ExcelConverter[List[Cell], HNil] =
    new ExcelConverter[List[Cell], HNil] {
      def from(cells: List[Cell]): Try[HNil] = cells match {
        case Nil =>
          Success(HNil)
        case _ => fail("Cannot convert '" ++ cells.toString ++ "' to HNil")
      }

      //      def to(n: HNil) = ""
    }

  implicit def deriveHCons[V, T <: HList]
  (implicit scv: Lazy[ExcelConverter[Cell, V]], sct: Lazy[ExcelConverter[List[Cell], T]])
  : ExcelConverter[List[Cell], V :: T] =
    new ExcelConverter[List[Cell], V :: T] {

      def from(r: List[Cell]): Try[V :: T] = r match {
        case Cons(before, after) =>
          for {
            front <- scv.value.from(before)
            back <- sct.value.from(after)
          } yield front :: back

        case _ => fail("Cannot convert '" ++ r.toString ++ "' to HList")
      }

      //      def to(ft: V :: T): String = {
      //        scv.value.to(ft.head) ++ "," ++ sct.value.to(ft.tail)
      //      }
    }
  implicit def deriveHConsOption[V, T <: HList]
  (implicit scv: Lazy[ExcelConverter[Cell, V]], sct: Lazy[ExcelConverter[List[Cell], T]])
  : ExcelConverter[List[Cell], Option[V] :: T] =
    new ExcelConverter[List[Cell], Option[V] :: T] {
      override def from(r: List[Cell]): Try[::[Option[V], T]] = {
        r match {
          case Cons(before, after) =>
            (for {
              front <- scv.value.from(before)
              back <- sct.value.from(after)
            } yield Some(front) :: back).orElse {
              sct.value.from(r).map(None :: _)
            }

          case List() => for {
            back <- sct.value.from(Nil)
          } yield None :: back

          case _ => fail("Cannot convert '" ++ r.toString ++ "' with Option to HList")
        }
      }
    }

  implicit def deriveHConsTry[V, T <: HList]
  (implicit scv: Lazy[ExcelConverter[Cell, V]], sct: Lazy[ExcelConverter[List[Cell], T]])
  : ExcelConverter[List[Cell], Try[V] :: T] =
    new ExcelConverter[List[Cell], Try[V] :: T] {
      override def from(r: List[Cell]): Try[::[Try[V], T]] = {
        r match {
          case Cons(before, after) =>
            val trial =for {
              front <- scv.value.from(before)
              back <- sct.value.from(after)
            } yield Success(front) :: back
              trial.orElse {
              sct.value.from(r).map(Failure(trial.failed.get) :: _)
            }

          case List() => for {
            back <- sct.value.from(Nil)
          } yield Failure(new NullPointerException) :: back

          case _ => fail("Cannot convert '" ++ r.toString ++ "' with Option to HList")
        }
      }
    }
  // Anything with a Generic
  implicit def deriveClass[A, R](implicit gen: Generic.Aux[A, R], conv: ExcelConverter[List[Cell], R])
  : ExcelConverter[List[Cell], A] = new ExcelConverter[List[Cell], A] {

    def from(row: List[Cell]): Try[A] = conv.from(row).map(gen.from)

    //    def to(a: A): String = conv.to(gen.to(a))
  }
}