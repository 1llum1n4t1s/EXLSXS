using EXLSXS.API.Services;
using Microsoft.AspNetCore.Mvc;

namespace EXLSXS.API.Controllers;

[ApiController]
[Route("api/[controller]")]
public class FinishController : ControllerBase
{
    private readonly IFontService _fontService;

    public FinishController(IFontService fontService)
    {
        _fontService = fontService;
    }

    [HttpGet("fonts")]
    public ActionResult<IEnumerable<string>> GetFonts()
    {
        return Ok(_fontService.GetInstalledFonts());
    }
}
