using Microsoft.AnalysisServices.AdomdClient;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using System.Data;

using System.Web.Http.Cors;
using API_REST_Cubo_Northwind.Models;

namespace API_REST_Cubo_Northwind.Controllers
{
    [EnableCors(origins:"*",headers:"*",methods:"*")]
    [RoutePrefix("v1/Analysis/Northwind")]
    public class NorthwindController : ApiController
    {
        [HttpGet]
        [Route("Testing")]
        public HttpResponseMessage Testing()
        {
            return Request.CreateResponse(HttpStatusCode.OK, "Prueba de API Exitosa");
        }

        [HttpGet]
        //[Route("Top5/{dim}/{order}")] string order = "DESC"
        [Route("Top5/{dim}")]
        public HttpResponseMessage Top5(string dim)
        {
            string dimension = string.Empty;

            switch (dim)
            {
                case "Cliente": dimension = "[Dim Cliente].[Dim Cliente Compania].CHILDREN";
                    break;
                case "Producto": dimension = "[Dim Producto].[Dim Producto Nombre].CHILDREN";
                    break;
                case "Empleado": dimension = "[Dim Empleado].[Employee Name].CHILDREN";
                    break;
                default:
                    dimension = "[Dim Cliente].[Dim Cliente Compania].CHILDREN";
                    break;
            }

            /*string WITH = @"
                WITH
                SET [TopVentas] AS
                NONEMPTY(
                    ORDER(
                        STRTOSET(@Dimension),
                        [Measures].[Fact Ventas Netas], " + order + @"
                    )
                )
            ";*/
            string WITH = @"
                WITH
                SET [TopVentas] AS
                NONEMPTY(
                    ORDER(
                        STRTOSET(@Dimension),
                        [Measures].[Fact Ventas Netas], DESC
                    )
                )
            ";
            string COLUMNS = @"
                NON EMPTY
                {
                    [Measures].[Fact Ventas Netas]
                }
                ON COLUMNS,    
            ";
            string ROWS = @"
                NON EMPTY
                {
                    HEAD([TopVentas], 5)
                }
                ON ROWS
            ";
            string CUBO_NAME = "[DWH Northwind]";
            string MDX_QUERY = WITH + @"SELECT " + COLUMNS + ROWS + "FROM" + CUBO_NAME;

            Debug.Write(MDX_QUERY);

            List<string> clients = new List<string>();
            List<decimal> ventas = new List<decimal>();
            List<dynamic> lstTabla = new List<dynamic>();

            dynamic result = new
            {
                datosDimension = clients,
                datosVenta = ventas,
                datosTabla = lstTabla
            };

            using (AdomdConnection cnn = new AdomdConnection(ConfigurationManager.ConnectionStrings["CuboNorhtwind"].ConnectionString))
            {
                cnn.Open();
                using (AdomdCommand cmd = new AdomdCommand(MDX_QUERY, cnn))
                {
                    cmd.Parameters.Add("Dimension", dimension);
                    using (AdomdDataReader dr = cmd.ExecuteReader(CommandBehavior.CloseConnection))
                    {
                        while (dr.Read())
                        {
                            clients.Add(dr.GetString(0));
                            ventas.Add(Math.Round(dr.GetDecimal(1)));

                            dynamic objTabla = new
                            {
                                descripcion = dr.GetString(0),
                                valor = Math.Round(dr.GetDecimal(1))
                            };

                            lstTabla.Add(objTabla);
                        }
                        dr.Close();
                    }
                }
            }
                return Request.CreateResponse(HttpStatusCode.OK, (object)result);
        }

        [HttpGet]
        //[Route("Top5/{dim}/{order}")] string order = "DESC"
        [Route("SerieH/{dim}")]
        public HttpResponseMessage SerieH(string dim)
        {
            string dimension = string.Empty;

            switch (dim)
            {
                case "Cliente":
                    dimension = "[Dim Cliente].[Dim Cliente Compania].CHILDREN";
                    break;
                case "Producto":
                    dimension = "[Dim Producto].[Dim Producto Nombre].CHILDREN";
                    break;
                case "Empleado":
                    dimension = "[Dim Empleado].[Employee Name].CHILDREN";
                    break;
                default:
                    dimension = "[Dim Cliente].[Dim Cliente Compania].CHILDREN";
                    break;
            }

            /*string WITH = @"
                WITH
                SET [TopVentas] AS
                NONEMPTY(
                    ORDER(
                        STRTOSET(@Dimension),
                        [Measures].[Fact Ventas Netas], " + order + @"
                    )
                )
            ";*/
            string WITH = @"
                WITH
                SET [TopVentas] AS
                NONEMPTY(
                    ORDER(
                        STRTOSET(@Dimension),
                        [Measures].[Fact Ventas Netas], DESC
                    )
                )
            ";
            string COLUMNS = @"
                NON EMPTY
                {
                    [Dim Tiempo].[Año].[Año]
                }
                ON COLUMNS,    
            ";
            string ROWS = @"
                NON EMPTY
                {
                    HEAD([TopVentas], 5)
                }
                ON ROWS
            ";
            string CUBO_NAME = "[DWH Northwind]";
            string MDX_QUERY = WITH + @"SELECT " + COLUMNS + ROWS + "FROM" + CUBO_NAME;

            Debug.Write(MDX_QUERY);

            List<string> clients = new List<string>();
            List<decimal> ventas = new List<decimal>();
            List<dynamic> lstTabla = new List<dynamic>();

            dynamic result = new
            {
                datosDimension = clients,
                //datosVenta = ventas,
                datosTabla = lstTabla
            };

            using (AdomdConnection cnn = new AdomdConnection(ConfigurationManager.ConnectionStrings["CuboNorhtwind"].ConnectionString))
            {
                cnn.Open();
                using (AdomdCommand cmd = new AdomdCommand(MDX_QUERY, cnn))
                {
                    cmd.Parameters.Add("Dimension", dimension);
                    using (AdomdDataReader dr = cmd.ExecuteReader(CommandBehavior.CloseConnection))
                    {
                        while (dr.Read())
                        {
                            //clients.Add(dr.GetName(1).Substring(21, 4));
                            ventas.Add(Math.Round(dr.GetDecimal(1)));
                            ventas.Add(Math.Round(dr.GetDecimal(2)));
                            ventas.Add(Math.Round(dr.GetDecimal(3)));

                            dynamic objTabla = new
                            {
                                data = ventas,
                                label = dr.GetString(0)
                            };

                            lstTabla.Add(objTabla);
                            ventas = new List<decimal>();
                        }
                        clients.Add(dr.GetName(1).Substring(21, 4));
                        clients.Add(dr.GetName(2).Substring(21, 4));
                        clients.Add(dr.GetName(3).Substring(21, 4));
                        dr.Close();
                    }
                }
            }
            return Request.CreateResponse(HttpStatusCode.OK, (object)result);
        }
        [HttpGet]
        //[Route("Top5/{dim}/{order}")] string order = "DESC"
        [Route("GetItemsByDimension/{dim}/{order}")]
        public HttpResponseMessage GetItemsByDimension(string dim, string order)
        {
            /*string WITH = @"
                WITH
                SET [TopVentas] AS
                NONEMPTY(
                    ORDER(
                        STRTOSET(@Dimension),
                        [Measures].[Fact Ventas Netas], " + order + @"
                    )
                )
            ";*/
            string WITH = @"
                WITH
                SET [OrderDimension] AS
                NONEMPTY(
                    ORDER(
                        {0}.CHILDREN,
                        {0}.CURRENTMEMBER.MEMBER_NAME, "+order+ @"
                    )
                )
            ";
            string COLUMNS = @"
                NON EMPTY
                {
                    [Measures].[Fact Ventas Netas]
                }
                ON COLUMNS,    
            ";
            string ROWS = @"
                NON EMPTY
                {
                    [OrderDimension]
                }
                ON ROWS
            ";
            string CUBO_NAME = "[DWH Northwind1]";
            WITH = string.Format(WITH, dim);
            string MDX_QUERY = WITH + @"SELECT " + COLUMNS + ROWS + "FROM" + CUBO_NAME;

            Debug.Write(MDX_QUERY);

            List<string> dimension = new List<string>();

            dynamic result = new
            {
                datosDimension = dimension
            };

            using (AdomdConnection cnn = new AdomdConnection(ConfigurationManager.ConnectionStrings["CuboNorhtwind2"].ConnectionString))
            {
                cnn.Open();
                using (AdomdCommand cmd = new AdomdCommand(MDX_QUERY, cnn))
                {
                    //cmd.Parameters.Add("Dimension", dimension);
                    using (AdomdDataReader dr = cmd.ExecuteReader(CommandBehavior.CloseConnection))
                    {
                        while (dr.Read())
                        {
                            dimension.Add(dr.GetString(0));
                        }
                        dr.Close();
                    }
                }
            }
            return Request.CreateResponse(HttpStatusCode.OK, (object)result);
        }

        [HttpPost]
        //[Route("Top5/{dim}/{order}")] string order = "DESC"
        [Route("GetDataPieByDimension/{dim}/{order}")]
        public HttpResponseMessage GetDataPieByDimension(string dim, string order, string [] values)
        {
            /*string WITH = @"
                WITH
                SET [TopVentas] AS
                NONEMPTY(
                    ORDER(
                        STRTOSET(@Dimension),
                        [Measures].[Fact Ventas Netas], " + order + @"
                    )
                )
            ";*/
            string WITH = @"
                WITH
                SET [OrderDimension] AS
                NONEMPTY(
                    ORDER(
                        STRTOSET(@Dimension),
                        [Measures].[Fact Ventas Netas], " + order + @"
                    )
                )
            ";
            string COLUMNS = @"
                NON EMPTY
                {
                    [Measures].[Fact Ventas Netas]
                }
                ON COLUMNS,    
            ";
            string ROWS = @"
                NON EMPTY
                {
                    [OrderDimension]
                }
                ON ROWS
            ";
            string CUBO_NAME = "[DWH Northwind]";
            //WITH = string.Format(WITH, dim);
            string MDX_QUERY = WITH + @"SELECT " + COLUMNS + ROWS + "FROM" + CUBO_NAME;

            Debug.Write(MDX_QUERY);

            List<string> dimension = new List<string>();
            List<decimal> ventas = new List<decimal>();
            List<dynamic> lstTabla = new List<dynamic>();

            dynamic result = new
            {
                datosDimension = dimension,
                datosVenta = ventas,
                datosTabla = lstTabla
            };

            string valoresDimension = string.Empty;
            foreach (var item in values)
            {
                valoresDimension += "{0}.[" + item + "],";
            }
            valoresDimension = valoresDimension.TrimEnd(',');
            valoresDimension = string.Format(valoresDimension, dim);
            valoresDimension = @"{" + valoresDimension + "}";

            using (AdomdConnection cnn = new AdomdConnection(ConfigurationManager.ConnectionStrings["CuboNorhtwind"].ConnectionString))
            {
                cnn.Open();
                using (AdomdCommand cmd = new AdomdCommand(MDX_QUERY, cnn))
                {
                    cmd.Parameters.Add("Dimension", valoresDimension);
                    using (AdomdDataReader dr = cmd.ExecuteReader(CommandBehavior.CloseConnection))
                    {
                        while (dr.Read())
                        {
                            dimension.Add(dr.GetString(0));
                            ventas.Add(Math.Round(dr.GetDecimal(1)));

                            dynamic objTabla = new
                            {
                                descripcion = dr.GetString(0),
                                valor = Math.Round(dr.GetDecimal(1))
                            };

                            lstTabla.Add(objTabla);
                        }
                        dr.Close();
                    }
                }
            }
            return Request.CreateResponse(HttpStatusCode.OK, (object)result);
        }

        [HttpPost]
        //[Route("Top5/{dim}/{order}")] string order = "DESC"
        [Route("SerieHDimension/{dim}/{dimYear}/{dimMes}/{order}")]
        public HttpResponseMessage SerieHDimension(string dim, string dimYear, string dimMes,  string order, [FromBody] Producto[] values)
        {

            /*string WITH = @"
                WITH
                SET [TopVentas] AS
                NONEMPTY(
                    ORDER(
                        STRTOSET(@Dimension),
                        [Measures].[Fact Ventas Netas], " + order + @"
                    )
                )
            ";*/
            string WITH = @"
                WITH
                SET [OrderDimension] AS
                NONEMPTY(
                    ORDER(
                        STRTOSET(@Dimension),
                        [Measures].[Fact Ventas Netas], " + order + @"
                    )
                )
            ";
            string COLUMNS = @"
                {
                    STRTOSET(@Meses)
                }
                ON COLUMNS,    
            ";
            string ROWS = @"
                NON EMPTY
                {
                    [OrderDimension]
                }
                ON ROWS
            ";
            string WHERE = @"
                WHERE STRTOSET(@Annios)
            ";
            string CUBO_NAME = "[DWH Northwind1]";
            string MDX_QUERY = WITH + @"SELECT " + COLUMNS + ROWS + "FROM" + CUBO_NAME+ WHERE;

            Debug.Write(MDX_QUERY);

            List<string> dimension = new List<string>();
            List<decimal> ventas = new List<decimal>();
            List<dynamic> lstTabla = new List<dynamic>();

            dynamic result = new
            {
                datosDimension = dimension,
                //datosVenta = ventas,
                datosTabla = lstTabla
            };

            string valoresFechas = string.Empty;
            string valoresAnnios = string.Empty;
            string valoresDimension = string.Empty;
            int meses = 0;

            foreach (var item in values)
            {
                for (int i = 0; i < item.Nombre.Length; i++)
                {
                    valoresDimension += "{0}.[" + item.Nombre[i] + "],";
                }
                for (int i = 0; i < item.Annio.Length; i++)
                {
                    valoresAnnios += "{0}.[" + item.Annio[i] + "],";
                }
                for (int i = 0; i < item.Mese.Length; i++)
                {
                    valoresFechas += "{0}.[" + item.Mese[i] + "],";
                    meses++;
                }
            }
            valoresDimension = valoresDimension.TrimEnd(',');
            valoresDimension = string.Format(valoresDimension, dim);
            valoresDimension = @"{" + valoresDimension + "}";

            valoresAnnios = valoresAnnios.TrimEnd(',');
            valoresFechas = valoresFechas.TrimEnd(',');
            valoresAnnios = string.Format(valoresAnnios, dimYear);
            valoresFechas = string.Format(valoresFechas, dimMes);
            valoresAnnios = @"{" + valoresAnnios + "}";
            valoresFechas = @"{" + valoresFechas + "}";

            using (AdomdConnection cnn = new AdomdConnection(ConfigurationManager.ConnectionStrings["CuboNorhtwind2"].ConnectionString))
            {
                cnn.Open();
                using (AdomdCommand cmd = new AdomdCommand(MDX_QUERY, cnn))
                {
                    cmd.Parameters.Add("Dimension", valoresDimension);
                    cmd.Parameters.Add("Meses", valoresFechas);
                    cmd.Parameters.Add("Annios", valoresAnnios);
                    Debug.Write(cmd);
                    using (AdomdDataReader dr = cmd.ExecuteReader(CommandBehavior.CloseConnection))
                    {
                        while (dr.Read())
                        {
                            for (int i = 1; i <= meses; i++)
                            {
                                try
                                {
                                    ventas.Add(Math.Round(dr.GetDecimal(i)));
                                }
                                catch
                                {

                                }
                                if (dimension.Count < meses)
                                {
                                    int encuentra = dr.GetName(i).IndexOf("&[");
                                    string mesDimension = dr.GetName(i).Substring(encuentra + 2);
                                    int encuentraFin = mesDimension.IndexOf("]");
                                    dimension.Add(mesDimension.Substring(0, encuentraFin));
                                }
                            }

                            dynamic objTabla = new
                            {
                                data = ventas,
                                label = dr.GetString(0)
                            };

                            lstTabla.Add(objTabla);
                            ventas = new List<decimal>();
                        }
                        dr.Close();
                    }
                }
            }
            return Request.CreateResponse(HttpStatusCode.OK, (object)result);
        }

        [HttpPost]
        //[Route("Top5/{dim}/{order}")] string order = "DESC"
        [Route("PieDimension/{dim}/{dimYear}/{dimMes}/{order}")]
        public HttpResponseMessage PieDimension(string dim, string dimYear, string dimMes, string order, [FromBody] Producto[] values)
        {

            /*string WITH = @"
                WITH
                SET [TopVentas] AS
                NONEMPTY(
                    ORDER(
                        STRTOSET(@Dimension),
                        [Measures].[Fact Ventas Netas], " + order + @"
                    )
                )
            ";*/
            string WITH = @"
                WITH
                SET [OrderDimension] AS
                NONEMPTY(
                    ORDER(
                        STRTOSET(@Dimension),
                        [Measures].[Fact Ventas Netas], " + order + @"
                    )
                )
            ";
            string COLUMNS = @"
                {
                    STRTOSET(@Annios)
                }
                ON COLUMNS,    
            ";
            string ROWS = @"
                NON EMPTY
                {
                    [OrderDimension]
                }
                ON ROWS
            ";
            string WHERE = @"
                WHERE STRTOSET(@Meses)
            ";
            string CUBO_NAME = "[DWH Northwind1]";
            string MDX_QUERY = WITH + @"SELECT " + COLUMNS + ROWS + "FROM" + CUBO_NAME + WHERE;

            Debug.Write(MDX_QUERY);

            List<string> dimension = new List<string>();
            List<decimal> ventas = new List<decimal>();
            List<dynamic> lstTabla = new List<dynamic>();

            dynamic result = new
            {
                datosDimension = dimension,
                datosVenta = ventas,
                datosTabla = lstTabla
            };

            string valoresFechas = string.Empty;
            string valoresAnnios = string.Empty;
            string valoresDimension = string.Empty;
            int meses = 0;

            foreach (var item in values)
            {
                for (int i = 0; i < item.Nombre.Length; i++)
                {
                    valoresDimension += "{0}.[" + item.Nombre[i] + "],";
                }
                for (int i = 0; i < item.Annio.Length; i++)
                {
                    valoresAnnios += "{0}.[" + item.Annio[i] + "],";
                }
                for (int i = 0; i < item.Mese.Length; i++)
                {
                    valoresFechas += "{0}.[" + item.Mese[i] + "],";
                    meses++;
                }
            }
            valoresDimension = valoresDimension.TrimEnd(',');
            valoresDimension = string.Format(valoresDimension, dim);
            valoresDimension = @"{" + valoresDimension + "}";

            valoresAnnios = valoresAnnios.TrimEnd(',');
            valoresFechas = valoresFechas.TrimEnd(',');
            valoresAnnios = string.Format(valoresAnnios, dimYear);
            valoresFechas = string.Format(valoresFechas, dimMes);
            valoresAnnios = @"{" + valoresAnnios + "}";
            valoresFechas = @"{" + valoresFechas + "}";

            using (AdomdConnection cnn = new AdomdConnection(ConfigurationManager.ConnectionStrings["CuboNorhtwind2"].ConnectionString))
            {
                cnn.Open();
                using (AdomdCommand cmd = new AdomdCommand(MDX_QUERY, cnn))
                {
                    cmd.Parameters.Add("Dimension", valoresDimension);
                    cmd.Parameters.Add("Meses", valoresFechas);
                    cmd.Parameters.Add("Annios", valoresAnnios);
                    Debug.Write(cmd);
                    using (AdomdDataReader dr = cmd.ExecuteReader(CommandBehavior.CloseConnection))
                    {
                        decimal total = 0;
                        while (dr.Read())
                        {
                            for (int i = 1; i <= meses; i++)
                            {
                                try
                                {
                                    total += Math.Round(dr.GetDecimal(i));
                                }
                                catch
                                {

                                }
                            }

                            dimension.Add(dr.GetString(0));
                            ventas.Add(total);

                            dynamic objTabla = new
                            {
                                descripcion = dr.GetString(0),
                                valor = total
                            };

                            lstTabla.Add(objTabla);
                            total = 0;
                        }
                        dr.Close();
                    }
                }
            }
            return Request.CreateResponse(HttpStatusCode.OK, (object)result);
        }

    }
}
