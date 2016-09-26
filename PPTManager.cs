using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Shell32;



namespace OfficeLib {
    public class PPTManager {

        string pptFile = "";

        //パワーポイント用
        Microsoft.Office.Interop.PowerPoint.Application app = null;
        Microsoft.Office.Interop.PowerPoint.Presentation ppt = null;

        
     
        //SlideShowNextSlide;

        string output_dir = "C:\\SlideJPEG";
        string pptName;
        string pptCreateTime;
        string stParentName;

        PPTInfo pptInfo;

        int count;


        Action<int> _slideNextMethod;
        public Action<int> slideNextMethod {
            set { _slideNextMethod = value; }
        }



        public PPTManager(string pptFile) {


            //  Directory.Delete(@"C:\SlideJPEG");
            Directory.CreateDirectory(@"C:\SlideJPEG");


            stParentName = Path.GetDirectoryName(pptFile);
            //  pptInfo.pptSingleFilename = Path.GetFileName(pptFile);

            // PPTのインスタンス作成  
            app = new Microsoft.Office.Interop.PowerPoint.Application();
            // オープン  
            ppt = app.Presentations.Open(pptFile
                , Microsoft.Office.Core.MsoTriState.msoTrue
                , Microsoft.Office.Core.MsoTriState.msoFalse
                , Microsoft.Office.Core.MsoTriState.msoFalse
            );
            count = ppt.Slides.Count;

            pptInfo = new PPTInfo(count);
            pptInfo.fullFilename = pptFile;
            pptInfo.singleFilename = Path.GetFileNameWithoutExtension(pptFile);
            pptInfo.createTime = File.GetCreationTime(pptFile).ToString();

            //追加
            stParentName = Path.GetDirectoryName(pptFile);
            pptName = Path.GetFileName(pptFile);
            ShellClass shell = new ShellClass();
            Folder f = shell.NameSpace(stParentName);
            FolderItem item = f.ParseName(pptName);
            pptInfo.presenter = f.GetDetailsOf(item, 20);
            
        }



        /* *
         *  スライドのタイトルとノート情報を入手 
         */
        public PPTInfo getSlideInfo() {

            //各スライドの処理
            for (int i = 0; i < count; i++) {
                //ノート情報の入手
                pptInfo.note[i] = ppt.Slides[i + 1].NotesPage.Shapes.Placeholders[2].TextFrame.TextRange.Text;
                foreach (Microsoft.Office.Interop.PowerPoint.Shape shape in ppt.Slides[i + 1].Shapes) {
                    if (shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoFalse) {
                        continue;
                    }
                    //タイトルの入手
                    if (shape.Name == "Title 1") {
                        pptInfo.titles[i] = shape.TextFrame.TextRange.Text;
                        break;
                    }
                }



            }
            return pptInfo;

        }


        public void saveJPeg() {



            //  XMLManager xmlman = new XMLManager("c:\\temp\\test2.xml");
            int width = (int)ppt.PageSetup.SlideWidth;
            int height = (int)ppt.PageSetup.SlideHeight;



            // １ページずつ保存する-----------------      
            for (int i = 0; i < count; i++) {
                // JPEGとして保存
                pptInfo.jpgName[i] = output_dir + String.Format("\\slide{0:0000}.jpg", i);
                ppt.Slides[i + 1].Export(pptInfo.jpgName[i], "jpg", width, height);
            }


        }


        public void slideShow() {
            try {
             
                //イベントの設定
                app.SlideShowNextSlide += _slideNext;
                app.SlideShowBegin += _slideNext;

                // オープン  
                ppt = app.Presentations.Open(pptInfo.fullFilename,
                    Microsoft.Office.Core.MsoTriState.msoTrue,
                    Microsoft.Office.Core.MsoTriState.msoTrue,
                    Microsoft.Office.Core.MsoTriState.msoTrue);

                // スライドショーのページの指定  
                int[] SlideIdx = new int[ppt.Slides.Count];
                //number = SlideIdx.Length;//TODO
                for (int i = 0; i < SlideIdx.Length; i++) {
                    SlideIdx[i] = i + 1;
                }

                Microsoft.Office.Interop.PowerPoint.SlideRange range;
                range = ppt.Slides.Range(SlideIdx);
                range.SlideShowTransition.AdvanceOnTime = Microsoft.Office.Core.MsoTriState.msoTrue;
                range.SlideShowTransition.AdvanceTime = 500;

                // 設定  
                Microsoft.Office.Interop.PowerPoint.SlideShowSettings settings;
                settings = ppt.SlideShowSettings;

                settings.StartingSlide = 1;
                settings.EndingSlide = SlideIdx[SlideIdx.Length - 1];

                // スライドショーの開始  
                settings.Run();

                // 待機する  
                while (app.SlideShowWindows.Count >= 1) {
                    System.Threading.Thread.Sleep(100);
                }

            } finally {
                // 終了  
                if (ppt != null) {
                    ppt.Close();
                }

                // PPTを閉じる  
                if (app != null) {
                    app.Quit();
                }
            }

        }

        private void _slideNext(Microsoft.Office.Interop.PowerPoint.SlideShowWindow Wn) {


            //現在のスライド番号
            int currentNum = Wn.View.CurrentShowPosition;
            _slideNextMethod(currentNum);



        }


     

        public void close() {

            // 終了  
            if (ppt != null) {
                ppt.Close();
            }

            // PPTを閉じる  
            if (app != null) {
                app.Quit();
                app = null;
            }

        }


    
    
    }
}
