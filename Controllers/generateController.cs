using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using WebApplication3.views;
//using WebApplication3.views;

namespace WebApplication3.Controllers {
    [ApiController]
    [Route("/")]
    public class generateController : ControllerBase {
        [HttpGet]
        public string index() {
            return "Welcome to the pdf API";
        }
        [HttpPost("romaneio-clientes")]
        public FileStreamResult GenerateAsync([FromBody] reqModel request) {

            return new GeneratePdf().GeraPdf(request.placa, request.motorista, request.ajudante, request.dataSub, request.empresa, request.pedidos);
        }
        [HttpPost("romaneio-separacao-clientes")]
        public FileStreamResult GenerateRomaneioSeparaçãoClientes([FromBody] reqModel2 request) {
            Console.Write(request.fracionados);
            return new GeneratePdf().GeraPdf2(request.placa, request.motorista, request.ajudante, request.dataSub, request.dados);
        }
        [HttpPost("romaneio-separacao-nao-fracionados")]
        public FileStreamResult GenerateRomaneioSeparaçãoNaoFracionados([FromBody] reqModel2 request) {

            return new GeneratePdf(fracionados: request.fracionados).GeraPdf3(request.placa, request.motorista, request.ajudante, request.dataSub, request.dados);
        }
        [HttpPost("romaneio-separacao-fracionados")]
        public FileStreamResult GenerateRomaneioSeparaçãoFracionados([FromBody] reqModel2 request) {

            return new GeneratePdf(fracionados: request.fracionados).GeraPdf4(request.placa, request.motorista, request.ajudante, request.dataSub, request.dados);
        }
        [HttpPost("romaneio-separacao-saudali")]
        public FileStreamResult GenerateRomaneioSeparaçãoSaudali([FromBody] reqModel3 request) {

            return new GeneratePdf(fracionados: request.fracionados, congelados: request.congelados).GeraPdf5(request.placa, request.motorista, request.ajudante, request.dataSub, request.produtos);
        }
        [HttpPost("romaneio-separacao-congelados-saudali")]
        public FileStreamResult GenerateRomaneioSeparaçãoCongeladosSaudali([FromBody] reqModel3 request) {

            return new GeneratePdf(fracionados: request.fracionados, congelados: request.congelados).GeraPdf6(request.placa, request.motorista, request.ajudante, request.dataSub, request.produtos);
        }
        [HttpPost("romaneio-separacao-fracionados-saudali")]
        public FileStreamResult GenerateRomaneioSeparaçãoFracionadosSaudali([FromBody] reqModel3 request) {

            return new GeneratePdf(fracionados: request.fracionados, congelados: request.congelados).GeraPdf7(request.placa, request.motorista, request.ajudante, request.dataSub, request.produtos);
        }

        [HttpPost("romaneio-separacao-clientes-saudali")]
        public FileStreamResult GenerateRomaneioSeparaçãoClientesSaudali([FromBody] reqModel6 request) {

            return new GeneratePdf().GeraPdf11(request.placa, request.motorista, request.ajudante, request.dataSub, request.orders);
        }

        [HttpPost("romaneio-nota-saudali")]
        public FileStreamResult GenerateRomaneioNotaSaudali([FromBody] reqModel4 request) {

            return new GeneratePdf().GeraPdf8(
                request.documento, 
                request.cliente, 
                request.bairro, 
                request.cidade, 
                request.valor, 
                request.peso,
                request.volume,
                request.dataSub,
                request.motorista,
                request.ajudante,
                request.placa,
                request.produtos
           );
        }
        [HttpPost("romaneio-nota-frisa")]
        public FileStreamResult GenerateRomaneioNotaFrisa([FromBody] reqModel5 request) {

            return new GeneratePdf().GeraPdf9(
                request.documento,
                request.cliente,
                request.bairro,
                request.cidade,
                request.valor,
                request.peso,
                request.volume,
                request.dataSub,
                request.motorista,
                request.ajudante,
                request.placa,
                request.dados
           );
        }
    }
}
