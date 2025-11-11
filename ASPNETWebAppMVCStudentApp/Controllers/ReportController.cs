using System;
using System.Linq;
using System.Web.Mvc;
using OfficeOpenXml;
using ASPNETWebAppMVCStudentApp.Models; // Change this to match your project name

namespace ASPNETWebAppMVCStudentApp.Controllers // Change this to match your project name
{
    public class ReportController : Controller
    {
        private SchoolDBEntities db = new SchoolDBEntities();

        // GET: Report
        public ActionResult Index()
        {
            // Get all courses for the dropdown
            ViewBag.Courses = new SelectList(db.Courses, "CourseID", "Title");
            return View();
        }

        // POST: Report/GenerateReport
        [HttpPost]
        public ActionResult GenerateReport(int courseId)
        {
            // Get course information
            var course = db.Courses.Find(courseId);
            
            if (course == null)
            {
                return HttpNotFound();
            }

            // Get all students enrolled in this course
            var enrollments = db.Enrollments
                .Where(e => e.CourseID == courseId)
                .Select(e => new
                {
                    StudentID = e.StudentID,
                    FirstName = e.Students.FirstName,
                    LastName = e.Students.LastName,
                    Grade = e.Grade
                })
                .ToList();

            // Create Excel file
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Student Report");

                // Title
                worksheet.Cells[1, 1].Value = "Student Enrollment Report";
                worksheet.Cells[1, 1, 1, 4].Merge = true;
                worksheet.Cells[1, 1].Style.Font.Size = 16;
                worksheet.Cells[1, 1].Style.Font.Bold = true;

                // Course information
                worksheet.Cells[2, 1].Value = "Course:";
                worksheet.Cells[2, 2].Value = course.Title;
                worksheet.Cells[2, 1].Style.Font.Bold = true;

                // Headers (row 4)
                worksheet.Cells[4, 1].Value = "Student ID";
                worksheet.Cells[4, 2].Value = "First Name";
                worksheet.Cells[4, 3].Value = "Last Name";
                worksheet.Cells[4, 4].Value = "Grade";

                // Make headers bold
                worksheet.Cells[4, 1, 4, 4].Style.Font.Bold = true;

                // Add data starting from row 5
                int row = 5;
                foreach (var enrollment in enrollments)
                {
                    worksheet.Cells[row, 1].Value = enrollment.StudentID;
                    worksheet.Cells[row, 2].Value = enrollment.FirstName;
                    worksheet.Cells[row, 3].Value = enrollment.LastName;
                    worksheet.Cells[row, 4].Value = enrollment.Grade;
                    row++;
                }

                // Auto-fit columns
                worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

                // Add borders
                worksheet.Cells[4, 1, row - 1, 4].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                worksheet.Cells[4, 1, row - 1, 4].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                worksheet.Cells[4, 1, row - 1, 4].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                worksheet.Cells[4, 1, row - 1, 4].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

                // Generate file name
                string fileName = $"StudentReport_{course.Title.Replace(" ", "_")}_{DateTime.Now:yyyyMMdd}.xlsx";

                // Return the Excel file
                return File(
                    package.GetAsByteArray(),
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    fileName
                );
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}