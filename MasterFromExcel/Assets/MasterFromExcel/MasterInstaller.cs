/////////////////////////////////////////////////
//自動生成ファイルです！直接編集しないでください！//
/////////////////////////////////////////////////

using Zenject;

namespace Master
{
    public class MasterInstaller : MonoInstaller
    {
        public override void InstallBindings()
        {
            Container.Bind<IAaaaDao>().To<AaaaDao>().AsSingle();
        }
    }
}