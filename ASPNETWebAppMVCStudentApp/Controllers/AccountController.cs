using ASPNETWebAppMVCStudentApp.Models;
using System.Linq;
using System.Web.Mvc;
public class AccountController : Controller
{
    private SchoolDBEntities db = new SchoolDBEntities();

    // GET: Account/Login
    public ActionResult Login()
    {
        return View();
    }

    // POST: Account/Login
    [HttpPost]
    [ValidateAntiForgeryToken]
    public ActionResult Login(string username, string password)
    {
        var user = db.tblUser.FirstOrDefault(u => u.Username == username && u.Password == password);

        if (user != null)
        {
            Session["UserID"] = user.UserID;
            Session["Username"] = user.Username;
            return RedirectToAction("Index", "Home");
        }

        ViewBag.Error = "Invalid username or password";
        return View();
    }

    // GET: Account/Logout
    public ActionResult Logout()
    {
        Session.Clear();
        return RedirectToAction("Login");
    }
}