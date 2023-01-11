using Microsoft.AspNetCore.Mvc;

// For more information on enabling Web API for empty projects, visit https://go.microsoft.com/fwlink/?LinkID=397860

namespace ALFA.Controllers
{
    [Route("api/[controller]")]
    [ApiController]

    public class ConvercaoController : ControllerBase
    {
        // GET: api/<ConvercaoController>
        [HttpGet]
        public void Get() => Covercao.ConversaoXLSAsync();
    }
}
