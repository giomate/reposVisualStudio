﻿using System.Collections.Generic;
using Autodesk.Forge.DesignAutomation.Model;

namespace Interaction
{
    /// <summary>
    /// Customizable part of Publisher class.
    /// </summary>
    internal partial class Publisher
    {
        /// <summary>
        /// Constants.
        /// </summary>
        private static class Constants
        {
            private const int EngineVersion = 23;
            public static readonly string Engine = $"Autodesk.Inventor+{EngineVersion}";

            public const string Description = "PUT DESCRIPTION HERE";

            internal static class Bundle
            {
                public static readonly string Id = "DebugPluginLocally1";
                public const string Label = "alpha";

                public static readonly AppBundle Definition = new AppBundle
                {
                    Engine = Engine,
                    Id = Id,
                    Description = Description
                };
            }

            internal static class Activity
            {
                public static readonly string Id = Bundle.Id;
                public const string Label = Bundle.Label;
            }

            internal static class Parameters
            {
                public const string InventorDoc = nameof(InventorDoc);
                public const string OutputIpt = nameof(OutputIpt);
            }
        }


        /// <summary>
        /// Get command line for activity.
        /// </summary>
        private static List<string> GetActivityCommandLine()
        {
            return new List<string> { $"$(engine.path)\\InventorCoreConsole.exe /al $(appbundles[{Constants.Activity.Id}].path) /i $(args[{Constants.Parameters.InventorDoc}].path)" };
        }

        /// <summary>
        /// Get activity parameters.
        /// </summary>
        private static Dictionary<string, Parameter> GetActivityParams()
        {
            return new Dictionary<string, Parameter>
                    {
                        {
                            Constants.Parameters.InventorDoc,
                            new Parameter
                            {
                                Verb = Verb.Get,
                                Description = "IPT file to process"
                            }
                        },
                        {
                            Constants.Parameters.OutputIpt,
                            new Parameter
                            {
                                Verb = Verb.Put,
                                LocalName = "result.ipt",
                                Description = "Resulting IPT",
                                Ondemand = false,
                                Required = false
                            }
                        }
                    };
        }

        /// <summary>
        /// Get arguments for workitem.
        /// </summary>
        private static Dictionary<string, IArgument> GetWorkItemArgs()
        {
            // TODO: update the URLs below with real values
            return new Dictionary<string, IArgument>
                    {
                        {
                            Constants.Parameters.InventorDoc,
                            new XrefTreeArgument
                            {
                                Url = "!!! CHANGE ME !!!"
                            }
                        },
                        {
                            Constants.Parameters.OutputIpt,
                            new XrefTreeArgument
                            {
                                Verb = Verb.Put,
                                Url = "!!! CHANGE ME !!!"
                            }
                        }
                    };
        }
    }
}
