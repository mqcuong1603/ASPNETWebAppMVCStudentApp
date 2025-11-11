using System.Web;
using System.Web.Mvc;

public class AuthorizeUser : ActionFilterAttribute
{
    public override void OnActionExecuting(ActionExecutingContext filterContext)
    {
        if (HttpContext.Current.Session["UserID"] == null)
        {
            filterContext.Result = new RedirectResult("~/Account/Login");
        }
    }
}