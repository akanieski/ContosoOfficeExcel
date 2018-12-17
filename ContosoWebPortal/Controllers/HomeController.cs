using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using ContosoWebPortal.Models;
using System.IO;
using Microsoft.Graph;

namespace ContosoWebPortal.Controllers
{
    public static class ControllerExtensions {
        public static string GetCurrentUserId (this Controller c)
        {
            return c.User.Claims.Single(x => x.Type == "http://schemas.microsoft.com/identity/claims/objectidentifier").Value;
        }
    }
    [Authorize]
    public class HomeController : Controller
    {
        private IGraphService _graphService;
        public HomeController(IGraphService graphService)
        {
            _graphService = graphService;
        }



        [HttpGet("testing")]
        public async Task<IActionResult> GetDrive()
        {
            
            return Ok();
        }
        public IActionResult Index()
        {
            return View();
        }

        public IActionResult About()
        {
            ViewData["Message"] = "Your application description page.";

            return View();
        }

        public IActionResult Contact()
        {
            ViewData["Message"] = "Your contact page.";

            return View();
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [AllowAnonymous]
        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
