﻿using Models.DAO;
using Models.EF;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using WebsiteDungCuHocTap.Common;
using WebsiteDungCuHocTap.Models;

namespace WebsiteDungCuHocTap.Areas.Admin.Controllers
{
    public class ProductController : HomeController
    {
        DBShop db = new DBShop();

        [HasCredential(RoleId = "VIEW_PRODUCT")]
        public ActionResult Show()
        {
            var session = (UserLogin)Session[WebsiteDungCuHocTap.Common.Commoncontent.user_sesion_admin];
            ViewBag.username = session.Username;

            var productViewModels = (from a in db.Products
                         join b in db.Providers on a.ProviderId equals b.ProviderId
                         join c in db.Categories on a.CateId equals c.CategoryId
                         select new ProductViewModel
                         {
                             ProductId = a.ProductId,
                             Name = a.Name,
                             Description = a.Description,
                             Discount = a.Discount,
                             ProviderName = b.Name,
                             CateName = c.Name,
                             Price = a.Price,
                             Quantity = a.Quantity,
                             StartDate = a.StartDate,
                             EndDate = a.EndDate,
                             Photo = a.Photo,
                         }).ToList();

            return View(productViewModels);
        }

        [HttpGet]
        [HasCredential(RoleId = "ADD_PRODUCT")]
        public ActionResult Add()
        {
            var session = (UserLogin)Session[WebsiteDungCuHocTap.Common.Commoncontent.user_sesion_admin];
            ViewBag.username = session.Username;

            ViewBag.ListCate = new SelectList(db.Categories.ToList(), "CategoryId", "Name");
            ViewBag.ListProvider = new SelectList(db.Providers.ToList(), "ProviderId", "Name");
            return View();
        }

        [HttpPost]
        [HasCredential(RoleId = "ADD_PRODUCT")]
        public ActionResult Add(ProductViewModel n, HttpPostedFileBase UploadImage)
        {
            var session = (UserLogin)Session[WebsiteDungCuHocTap.Common.Commoncontent.user_sesion_admin];
            ViewBag.username = session.Username;

            ViewBag.ListCate = new SelectList(db.Categories.ToList(), "CategoryId", "Name");
            ViewBag.ListProvider = new SelectList(db.Providers.ToList(), "ProviderId", "Name");


            if (UploadImage == null || UploadImage.ContentLength == 0)
            {
                ModelState.AddModelError("Photo", "Vui lòng chọn ảnh sản phẩm.");
                return View(n);
            }

            if (db.Products.Any(p => p.ProductId == n.ProductId))
            {
                ModelState.AddModelError("ProductError", "Mã sản phẩm đã tồn tại.");
                return View(n);
            }
            if (n.Price <= 0)
            {
                ModelState.AddModelError("ProductError", "Giá tiền không được bé hơn hoặc bằng 0.");
                return View(n);
            }

            try
            {
                var fileName = Path.GetFileName(UploadImage.FileName);
                var path = Path.Combine(Server.MapPath("~/image"), fileName);
                UploadImage.SaveAs(path);
                n.Photo = fileName;

                var product = new Product
                {
                    ProductId = n.ProductId,
                    Name = n.Name,
                    Photo = n.Photo,
                    Price = n.Price,
                    Quantity = n.Quantity,
                    StartDate = n.StartDate,
                    EndDate = n.EndDate,
                    CateId = n.CateId,
                    Description = n.Description,
                    Discount = n.Discount ?? 0,
                    ProviderId = n.ProviderId
                };

                db.Products.Add(product);
                db.SaveChanges();

                return RedirectToAction("Show");
            }
            catch (Exception ex)
            {
                ModelState.AddModelError("", "Có lỗi xảy ra: " + ex.Message);
                return View(n);
            }
        }


        [HttpGet]
        [HasCredential(RoleId = "EDIT_PRODUCT")]
        public ActionResult Edit(int ProductId)
        {
            var session = (UserLogin)Session[WebsiteDungCuHocTap.Common.Commoncontent.user_sesion_admin];
            ViewBag.username = session.Username;

            var model = (from a in db.Products
                         join b in db.Providers on a.ProviderId equals b.ProviderId
                         join c in db.Categories on a.CateId equals c.CategoryId
                         where a.ProductId == ProductId
                         select new ProductViewModel
                         {
                             ProductId = a.ProductId,
                             Name = a.Name,
                             Description = a.Description,
                             Discount = a.Discount,
                             ProviderName = b.Name,
                             CateName = c.Name,
                             Price = a.Price,
                             Quantity = a.Quantity,
                             StartDate = a.StartDate,
                             EndDate = a.EndDate,
                             Photo = a.Photo,
                             CateId = a.CateId
                         }).ToList();

            ViewBag.ListCate = new SelectList(db.Categories.ToList(), "CategoryId", "Name");
            ViewBag.ListProvider = new SelectList(db.Providers.ToList(), "ProviderId", "Name");
            var models = model.Where(n => n.ProductId == ProductId).First();
            return View(models);
        }

        [HttpPost]
        [HasCredential(RoleId = "EDIT_PRODUCT")]
        public ActionResult Edit(ProductViewModel n, HttpPostedFileBase UploadImage)

        {
            var session = (UserLogin)Session[WebsiteDungCuHocTap.Common.Commoncontent.user_sesion_admin];
            ViewBag.username = session.Username;

            ViewBag.ListCate = new SelectList(db.Categories.ToList(), "CategoryId", "Name");
            ViewBag.ListProvider = new SelectList(db.Providers.ToList(), "ProviderId", "Name");
            if (ModelState.IsValid)
            {
                ProductDao a = new ProductDao();
                if (UploadImage != null)
                {
                    // Delete exiting file
                    //System.IO.File.Delete(Path.Combine(Server.MapPath("~/image"), n.Photo));
                    // Save new file
                    string fileName = Path.GetFileName(UploadImage.FileName);
                    string path = Path.Combine(Server.MapPath("~/image"), fileName);
                    UploadImage.SaveAs(path);
                    n.Photo = fileName;

                }
                var model = db.Products.FirstOrDefault(m => m.ProductId == n.ProductId);
                model.ProductId = n.ProductId;
                model.Name = n.Name;
                model.Photo = n.Photo;
                model.Price = n.Price;
                model.Quantity = n.Quantity;
                model.StartDate = n.StartDate;
                model.EndDate = n.EndDate;
                model.CateId = model.CateId;
                model.ProductId = n.ProductId;
                model.Description = n.Description;
                model.Discount = n.Discount;
                model.ProviderId = n.ProviderId;
                db.SaveChanges();
                return RedirectToAction("Show", new { CateId = n.CateId });
            }
            else
            {
                ModelState.AddModelError("", "Ngày kết thúc phải muộn hơn ngày bắt đầu");
                return View();
            }
        }

        //[HttpGet]
        //[HasCredential(RoleId = "DELETE_PRODUCT")]
        //public ActionResult Delete()
        //{
        //    var session = (UserLogin)Session[WebsiteDungCuHocTap.Common.Commoncontent.user_sesion_admin];
        //    ViewBag.username = session.Username;
        //    return View();
        //}
        [HttpPost]
        [HasCredential(RoleId = "DELETE_PRODUCT")]
        public ActionResult Delete(FormCollection formCollection)
        {
            string[] ids = formCollection["ProductId"].Split(new char[] { ',' });

         

            foreach (string id in ids)
            {
                var model = db.Products.Find(Convert.ToInt32(id));
                var orderDetails = db.OrderDetails.Where(od => od.ProductId == model.ProductId).ToList();
                db.OrderDetails.RemoveRange(orderDetails);
                db.Products.Remove(model);
                db.SaveChanges();
            }
            return RedirectToAction("Show");
        }

        public ActionResult Menu()
        {
            var session = (UserLogin)Session[WebsiteDungCuHocTap.Common.Commoncontent.user_sesion_admin];
            ViewBag.username = session.Username;

            var model = new CategoryDao().ListCategory();
            return PartialView(model);
        }
    }
}