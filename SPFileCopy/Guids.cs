// Guids.cs
// MUST match guids.h
using System;

namespace Britehouse.SPFileCopy
{
    static class GuidList
    {
        public const string guidSPFileCopyPkgString = "13250768-13a6-4632-8d49-a47011d5087c";
        public const string guidSPFileCopyCmdSetString = "8556d950-3dba-4c0e-89a6-42a8a0285851";

        public static readonly Guid guidSPFileCopyCmdSet = new Guid(guidSPFileCopyCmdSetString);
    };
}