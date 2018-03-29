using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using WordTemplate.Models;
using Xceed.Words.NET;

namespace WordTemplate.Controllers
{
    public class PeopleController : Controller
    {
        private readonly PersonContext _context;

        public PeopleController(PersonContext context)
        {
            _context = context;
        }

        // GET: People
        public async Task<IActionResult> Index()
        {
            return View(await _context.Person.ToListAsync());
        }

        // GET: People/Details/5
        public async Task<IActionResult> Details(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var person = await _context.Person
                .SingleOrDefaultAsync(m => m.Id == id);
            if (person == null)
            {
                return NotFound();
            }

            return View(person);
        }

        // GET: People/Create
        public IActionResult Create()
        {
            return View();
        }

        // POST: People/Create
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Create([Bind("Id,FirstName,LastName,Age,Email")] Person person)
        {
            if (ModelState.IsValid)
            {
                _context.Add(person);
                await _context.SaveChangesAsync();
                return RedirectToAction(nameof(Index));
            }
            return View(person);
        }

        private Person _CurrentPerson = new Person();
        private void DotheWordThings(Person person)
        {
            _CurrentPerson = person;
            using (DocX document = DocX.Load(@"Documents\PersonTemplate.docx"))
            {
                if (document.FindUniqueByPattern(@"<[\w \=]{3,}>", RegexOptions.IgnoreCase).Count ==
                   4)
                {
                    for (int i = 0; i < 4; i++)
                    {
                        document.ReplaceText(@"<(.*?)>", RegexMatchHandler, false, RegexOptions.IgnoreCase,
                           null, new Formatting());
                    }
                    document.SaveAs("Result.Docx");

                }
            }
        }

        public string RegexMatchHandler(string findStr)
        {
            switch (findStr)
            {
                case "FirstName":
                    return _CurrentPerson.FirstName;
                    break;
                case "LastName":
                    return _CurrentPerson.LastName;
                    break;
                case "Age":
                    return _CurrentPerson.Age.ToString();
                    break;
                case "Email":
                    return _CurrentPerson.Email;
                    break;

                default:
                    return "Could'nt Be Found";
            }
        }

        // GET: People/Edit/5
        public async Task<IActionResult> Edit(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var person = await _context.Person.SingleOrDefaultAsync(m => m.Id == id);
            if (person == null)
            {
                return NotFound();
            }
            return View(person);
        }

        // POST: People/Edit/5
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Edit(int id, [Bind("Id,FirstName,LastName,Age,Email")] Person person)
        {
            if (id != person.Id)
            {
                return NotFound();
            }

            if (ModelState.IsValid)
            {
                try
                {
                    _context.Update(person);
                    DotheWordThings(person);
                    await _context.SaveChangesAsync();
                }
                catch (DbUpdateConcurrencyException)
                {
                    if (!PersonExists(person.Id))
                    {
                        return NotFound();
                    }
                    else
                    {
                        throw;
                    }
                }
                return RedirectToAction(nameof(Index));
            }
            return View(person);
        }

        // GET: People/Delete/5
        public async Task<IActionResult> Delete(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var person = await _context.Person
                .SingleOrDefaultAsync(m => m.Id == id);
            if (person == null)
            {
                return NotFound();
            }

            return View(person);
        }

        // POST: People/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> DeleteConfirmed(int id)
        {
            var person = await _context.Person.SingleOrDefaultAsync(m => m.Id == id);
            _context.Person.Remove(person);
            await _context.SaveChangesAsync();
            return RedirectToAction(nameof(Index));
        }

        private bool PersonExists(int id)
        {
            return _context.Person.Any(e => e.Id == id);
        }

        public async Task<ActionResult> DownloadTemplate(int id)
        {
            var person = await _context.Person.SingleOrDefaultAsync(m => m.Id == id);
            if (person != null)
            {
                 DotheWordThings(person);
            return  Download("Result.docx");
            }
            return RedirectToAction(nameof(Index));
        }

        public FileResult Download(string filename)
        {
            byte[] fileBytes = System.IO.File.ReadAllBytes(filename);
            string fileName = "Profile.docx";
            return File(fileBytes, System.Net.Mime.MediaTypeNames.Application.Octet, fileName);
        }
    }
}
