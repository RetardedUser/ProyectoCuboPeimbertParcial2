﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Http;

using Newtonsoft.Json.Serialization;
using System.Web.Http.Cors;
using Newtonsoft.Json;
using System.Web.Cors;

namespace API_REST_Cubo_Northwind
{
    public static class WebApiConfig
    {
        public static void Register(HttpConfiguration config)
        {
            // Configuración y servicios de API web
            var formatters = GlobalConfiguration.Configuration.Formatters;
            var jsonFormatters = formatters.JsonFormatter;
            var settings = jsonFormatters.SerializerSettings;

            jsonFormatters.SerializerSettings.PreserveReferencesHandling = PreserveReferencesHandling.None;
            config.Formatters.Remove(config.Formatters.XmlFormatter);
            settings.Formatting = Formatting.Indented;
            settings.NullValueHandling = NullValueHandling.Ignore;
            settings.ContractResolver = new CamelCasePropertyNamesContractResolver();
            config.EnableCors();

            // Rutas de API web
            config.MapHttpAttributeRoutes();

            config.Routes.MapHttpRoute(
                name: "DefaultApi",
                routeTemplate: "api/{controller}/{id}",
                defaults: new { id = RouteParameter.Optional }
            );
        }
    }
}
