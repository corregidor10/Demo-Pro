using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace DemoProvidersWeb.Controllers
{
    public class HomeController : Controller
    {
        [SharePointContextFilter]
        public ActionResult Index()
        {
            User spUser = null;

            var spContext = SharePointContextProvider.Current.GetSharePointContext(HttpContext);

            var data = new List<TelefonoViewModel>();

            using (var clientContext = spContext.CreateUserClientContextForSPHost())
            {
                if (clientContext != null)
                {
                    var telefonosList = clientContext.Web.Lists.GetByTitle("Telefonos");

                    clientContext.Load(telefonosList);

                    clientContext.ExecuteQuery();

                    var query = new CamlQuery();
                    var telefonosItem = telefonosList.GetItems(query);

                    clientContext.Load(telefonosItem);
                    clientContext.ExecuteQuery();


                    foreach (var telItems in telefonosItem)
                    {
                        data.Add(TelefonoViewModel.FromListItem(telItems));
                    }


                    //    spUser = clientContext.Web.CurrentUser;

                    //    clientContext.Load(spUser, user => user.Title);

                    //    clientContext.ExecuteQuery();

                    //    ViewBag.UserName = spUser.Title;
                }
            }

            return View(data);
        }

        public ActionResult Add()
        {

            return View(new TelefonoViewModel());
        }
        [HttpPost]
        public ActionResult Add(TelefonoViewModel model)
        {

            var spContext = SharePointContextProvider.Current.GetSharePointContext(HttpContext);


            using (var clientContext = spContext.CreateUserClientContextForSPHost())
            {
                if (clientContext != null)
                {
                    var telefonosList = clientContext.Web.Lists.GetByTitle("Telefonos");

                    clientContext.Load(telefonosList);

                    clientContext.ExecuteQuery();

                    var listCreationInfo = new ListItemCreationInformation();

                    var item = telefonosList.AddItem(listCreationInfo);

                    item["Title"] = model.Nombre;
                    item["Numero"] = model.Numero;

                    item.Update();

                    clientContext.ExecuteQuery();


                }

                return RedirectToAction("Index", new { SPHostUrl = SharePointContext.GetSPHostUrl(HttpContext.Request).AbsoluteUri });// siempre hay que indicar el valor del SPHostUrl 


            };

        }


        public ActionResult Delete(int id)
        {


            var spContext = SharePointContextProvider.Current.GetSharePointContext(HttpContext);


            using (var clientContext = spContext.CreateUserClientContextForSPHost())
            {
                if (clientContext != null)
                {
                    var telefonosList = clientContext.Web.Lists.GetByTitle("Telefonos");

                    var telefonosItem = telefonosList.GetItemById(id);

                    telefonosItem.DeleteObject();

                    clientContext.ExecuteQuery();





                }


                return RedirectToAction("Index", new { SPHostUrl = SharePointContext.GetSPHostUrl(HttpContext.Request).AbsoluteUri });// siempre hay que indicar el valor del SPHostUrl 

            }
        }


        public ActionResult Update(int id)
        {

            var spContext = SharePointContextProvider.Current.GetSharePointContext(HttpContext);

            TelefonoViewModel model = null;


            using (var clientContext = spContext.CreateUserClientContextForSPHost())
            {
                if (clientContext != null)
                {
                    var telefonosList = clientContext.Web.Lists.GetByTitle("Telefonos");
                    clientContext.Load(telefonosList);


                    var telefonosItem = telefonosList.GetItemById(id);
                    clientContext.Load(telefonosItem);

                    clientContext.ExecuteQuery();


                    model = TelefonoViewModel.FromListItem(telefonosItem);






                }
            }



            return View(model);
        }

        [HttpPost]
        public ActionResult Update(TelefonoViewModel model)
        {


            var spContext = SharePointContextProvider.Current.GetSharePointContext(HttpContext);


            using (var clientContext = spContext.CreateUserClientContextForSPHost())
            {
                if (clientContext != null)
                {
                    var telefonosList = clientContext.Web.Lists.GetByTitle("Telefonos");

                    var telefonosItem = telefonosList.GetItemById(model.Id);


                    telefonosItem["Title"] = model.Nombre;
                    telefonosItem["Numero"] = model.Numero;


                    telefonosItem.Update();

                    clientContext.ExecuteQuery();



                }


                return RedirectToAction("Index", new { SPHostUrl = SharePointContext.GetSPHostUrl(HttpContext.Request).AbsoluteUri });// siempre hay que indicar el valor del SPHostUrl 

            }
        }



        // FIN CLASS
    }

}





