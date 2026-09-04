//#define ENGINE
#if ENGINE
using OpenDSSengine;
#else
using dss_sharp;
#endif

namespace ExecutorOpenDSS.Engine
{
    public static class EngineConfig
    {
#if ENGINE
        public const bool OpenDSSengine = true;
#else
        public const bool OpenDSSengine = false;
#endif
    }
}
// if (EngineConfig.OpenDSSengine)
