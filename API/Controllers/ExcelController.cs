using API.Models;
using Microsoft.AspNetCore.Mvc;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace API.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ExcelController : ControllerBase
    {

        [Route("Import")]
        [HttpPost]
        public IActionResult Import([FromForm] ImportExcelDTO request)
        {
            if (Path.GetExtension(request.File.FileName) != ".xls" && Path.GetExtension(request.File.FileName) != ".xlsx")
            {
                return BadRequest("Solo se admiten archivos excel");
            }

            var result = ReadExcelFile(request.File);

            return Ok(result);
        }

        [Route("Export")]
        [HttpPost]
        public IActionResult Export([FromBody] ExportExcelDTO request)
        {
            var employees = GetEmployees(request.DepartmentId);

            var bytes = CreateExcelFile(employees);

            return File(bytes, "application/vnd.ms-excel", "Employees.xlsx");
        }


        private static List<Employee> ReadExcelFile(IFormFile file)
        {
            var employees = new List<Employee>();
            int numberOfRecords = 0;

            IWorkbook workbook = null;

            if (Path.GetExtension(file.FileName) == ".xlsx")
            {
                workbook = new XSSFWorkbook(file.OpenReadStream());
            }
            else if (Path.GetExtension(file.FileName) == ".xls")
            {
                workbook = new HSSFWorkbook(file.OpenReadStream());
            }

            //Obtener la primera hoja del libro de excel
            ISheet sheet = workbook.GetSheetAt(0);

            int nameColumnIndex = 1;
            int lastNameColumnIndex = 2;
            int salaryColumnIndex = 6;

            //Nombre de Cabeceras - Fila 1
            string nameHeaderName = "Nombre";
            string lastNameHeaderName = "Apellido";
            string salaryHeaderName = "Salario";

            //Nombre de Cabeceras dinámicas
            if (sheet.GetRow(0) != null)
            {
                nameHeaderName = sheet.GetRow(0).GetCell(nameColumnIndex).StringCellValue;
                lastNameHeaderName = sheet.GetRow(0).GetCell(lastNameColumnIndex).StringCellValue;
                salaryHeaderName = sheet.GetRow(0).GetCell(salaryColumnIndex).StringCellValue;
            }

            for (int rowIndex = 1; rowIndex <= sheet.LastRowNum; rowIndex++)
            {
                //Verifica si la fila tiene datos
                if (sheet.GetRow(rowIndex) != null && !string.IsNullOrEmpty(sheet.GetRow(rowIndex).GetCell(nameColumnIndex).StringCellValue))
                {
                    numberOfRecords++;

                    string nameCellValue = sheet.GetRow(rowIndex).GetCell(nameColumnIndex).StringCellValue.Trim().ToUpperInvariant();
                    string lastNameCellValue = sheet.GetRow(rowIndex).GetCell(lastNameColumnIndex).StringCellValue;
                    decimal salaryCellValue = (decimal)sheet.GetRow(rowIndex).GetCell(salaryColumnIndex).NumericCellValue;

                    employees.Add(new Employee
                    {
                        Name = nameCellValue,
                        LastName = lastNameCellValue,
                        Salary = salaryCellValue
                    });

                }
            }

            return employees;
        }

        private static List<Employee> GetEmployees(int departmentId)
        {
            var employees = new List<Employee>() {
            new Employee{
                Id= 100,
                Name = "Steven",
                LastName= "King",
                Email= "sking@gmail.com",
                PhoneNumber = "(+54) 9 515-123-4567",
                HireDate = new DateTime(1987, 06, 17, 16, 10, 12),
                Salary = 24563.264M,
                DepartmentId = 90,
                IsMellennial = true
            },
            new Employee{
                Id= 103,
                Name = "Alexander",
                LastName= "Hunold",
                Email= "ahunold@gmail.com",
                PhoneNumber = "(+54) 9 590-423-5567",
                HireDate = new DateTime(2021, 12, 31, 7, 05, 22),
                Salary = 1999.98M,
                DepartmentId = 90,
                IsMellennial = false
            },
            new Employee{
                Id= 129,
                Name = "Laura",
                LastName= "Bissot",
                Email= "lbissot@gmail.com",
                PhoneNumber = "(+54) 9 650-121-2034",
                HireDate = new DateTime(1997, 03, 29, 21, 09, 0),
                Salary = 1800,
                DepartmentId = 50,
                IsMellennial = true
            }
            };

            if (departmentId == 0)
            {
                return employees;
            }
            else
            {
                return employees.Where(x => x.DepartmentId == departmentId).ToList();
            }

        }

        private static byte[] CreateExcelFile(List<Employee> employees)
        {
            IWorkbook workbook = new XSSFWorkbook();

            //Styles
            var headerStyle = workbook.CreateCellStyle();
            headerStyle.FillForegroundColor = HSSFColor.Yellow.Index2;
            headerStyle.FillPattern = FillPattern.SolidForeground;

            var dataFormat = workbook.CreateDataFormat();
            var dateTimeStyle = workbook.CreateCellStyle();
            dateTimeStyle.DataFormat = dataFormat.GetFormat("dd/MM/yyyy HH:mm:ss");

            //Colocar nombre de la pestaña
            ISheet worksheet = workbook.CreateSheet("Pestaña Nº1");

            int rowNumber = 0;
            IRow row = worksheet.CreateRow(rowNumber++);

            //Table Header
            ICell cell = row.CreateCell(0);
            cell.CellStyle = headerStyle;
            cell.SetCellValue("Id");

            cell = row.CreateCell(1);
            cell.CellStyle = headerStyle;
            cell.SetCellValue("Nombre");

            cell = row.CreateCell(2);
            cell.CellStyle = headerStyle;
            cell.SetCellValue("Apellido");

            cell = row.CreateCell(3);
            cell.CellStyle = headerStyle;
            cell.SetCellValue("Email");

            cell = row.CreateCell(4);
            cell.CellStyle = headerStyle;
            cell.SetCellValue("Teléfono");

            cell = row.CreateCell(5);
            cell.CellStyle = headerStyle;
            cell.SetCellValue("Fecha de contratación");

            cell = row.CreateCell(6);
            cell.CellStyle = headerStyle;
            cell.SetCellValue("Salario");

            cell = row.CreateCell(7);
            cell.CellStyle = headerStyle;
            cell.SetCellValue("Departamento");

            cell = row.CreateCell(8);
            cell.CellStyle = headerStyle;
            cell.SetCellValue("Es millennial");

            //Table Body
            foreach (var employee in employees)
            {
                row = worksheet.CreateRow(rowNumber++);

                //Table Header
                cell = row.CreateCell(0);
                cell.SetCellValue(employee.Id);

                cell = row.CreateCell(1);
                cell.SetCellValue(employee.Name);

                cell = row.CreateCell(2);
                cell.SetCellValue(employee.LastName);

                cell = row.CreateCell(3);
                cell.SetCellValue(employee.Email);

                cell = row.CreateCell(4);
                cell.SetCellValue(employee.PhoneNumber);

                cell = row.CreateCell(5);
                cell.CellStyle = dateTimeStyle;
                cell.SetCellValue(employee.HireDate);

                cell = row.CreateCell(6);
                cell.SetCellValue((double)employee.Salary);

                cell = row.CreateCell(7);
                cell.SetCellValue(employee.DepartmentId);

                cell = row.CreateCell(8);
                cell.SetCellValue(employee.IsMellennial ? "Y" : "N");
            }

            var ms = new MemoryStream();
            workbook.Write(ms, true);
            byte[] bytes = ms.ToArray();
            ms.Close();

            return bytes;
        }
    }
}
