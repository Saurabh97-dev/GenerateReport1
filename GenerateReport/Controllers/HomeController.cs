using System.Diagnostics;
using GenerateReport.Models;
using Microsoft.AspNetCore.Mvc;
using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Element;
using Microsoft.AspNetCore.Http;
using OfficeOpenXml;
using System.Collections.Generic;
using System.IO;
using Microsoft.AspNetCore.Hosting.Server;
using iText.Layout.Properties;

namespace GenerateReport.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index()
        {
            var data = new[]
            {
                new { SrNo = 1, Location = "Mohana", SSTNames = "Rakesh kumar Gupta\nG.K Shrivastav\nOmprakash Kane", Mobile = "9926225166\n9977021658\n9926245820", CheckingTime = "24 hours", Remark1 = "जी के श्रीवास्तव जी के द्वारा बताया गया है की चेकिंग बराबर की जा रही है। (11:42PM)", Remark2 = "जी के श्रीवास्तव जी के द्वारा बताया गया है की चेकिंग बराबर की जा रही है। (01:41AM)", Remark3 = "जी के श्रीवास्तव जी के द्वारा फोर नहीं उठाया गया है। (03:40AM)", Remark4 = "जी के श्रीवास्तव जी के द्वारा बताया गया है की चेकिंग बराबर की जा रही है। (05:42AM)" },
                new { SrNo = 2, Location = "Gohinda", SSTNames = "Piyush Chand\nRK Kokwani\nKapil Kumar", Mobile = "8124484832\n6232641721\n7906394885", CheckingTime = "", Remark1 = "पीयूष चंदजी जी के द्वारा बताया गया है कि चैकिंग बराबर हो रही है।  (11:43 pm)", Remark2 = "पीयूष चंदजी जी के द्वारा बताया गया है कि चैकिंग बराबर हो रही है।  (01:43 Am)", Remark3 = "पीयूष चंदजी जी के द्वारा बताया गया है कि चैकिंग बराबर हो रही है।  (03:42 Am)", Remark4 = "पीयूष चंदजी जी के द्वारा बताया गया है कि चैकिंग बराबर हो रही है।  (05:43 Am)" },
                new { SrNo = 3, Location = "Lodi Tiraha", SSTNames = "Lawrence kumar bodh\nM.C Sharma\nAjay Kumar", Mobile = "9993277488\n9425619912\n9406972494", CheckingTime = "", Remark1 = "एमसी शर्माजी के द्वारा बताया गया है कि चैकिंग बराबर हो रही है।  (11:46 pm)", Remark2 = "एमसी शर्माजी के द्वारा बताया गया है कि चैकिंग बराबर हो रही है।  (01:46 Am)", Remark3 = "ऍम सी शर्मा जी के द्वारा फोर नहीं उठाया गया है।  (03:46 Am)", Remark4 = "ऍम सी शर्मा जी के बताया गया है कि चैकिंग बराबर हो रही है।  (05:46 Am)" },
                new { SrNo = 4, Location = "Ginjora", SSTNames = "Ramesvar singh Alapuriya\nSunil Saxena\nJaynaryan Rajoriya", Mobile = "9926228570\n9826565033\n9926730423", CheckingTime = "", Remark1 = "जयनारायण राजोरिया जी के द्वारा बताया गया है कि चैकिंग बराबर हो रही है।(11:48 pm)", Remark2 = "जयनारायण राजोरिया जी के द्वारा बताया गया है कि चैकिंग बराबर हो रही है।(01:47 Am)", Remark3 = "जयनारायण राजोरिया जी के द्वारा बताया गया है कि चैकिंग बराबर हो रही है।(03:47 Am)", Remark4 = "जयनारायण राजोरिया जी के द्वारा बताया गया है कि चैकिंग बराबर हो रही है।(05:47 Am)" },
                new { SrNo = 5, Location = "Lidhora", SSTNames = "Shubham Kumar Gupta\nVijay Kumar Sharma\nBhola Gurjar", Mobile = "8120995050\n9753139003\n9630757536", CheckingTime = "", Remark1 = "विजय कुमार शर्मा जी के द्वारा बताया गया हे कि यह पे चेकिंग बराबर हो रही है (11:50pm)", Remark2 = "विजय कुमार शर्मा जी के द्वारा बताया गया हे कि यह पे चेकिंग बराबर हो रही है (01:50Am)", Remark3 = "विजय कुमार शर्मा जी के द्वारा फ़ोन नहीं उठाया गया है (03:50Am)", Remark4 = "विजय कुमार शर्मा जी के द्वारा फ़ोन नहीं उठाया गया है (05:50Am)" },
                new { SrNo = 6, Location = "Chandpura", SSTNames = "Padmakar Mokhribale\nNathuram Mahor\nSatyprakash Samoliya", Mobile = "9826240623\n9826280044\n9893601369", CheckingTime = "", Remark1 = "नाथूराम माहोर जी के द्वारा बताया गया हे कि यह पे चेकिंग बराबर हो रही है (11:53 pm)", Remark2 = "नाथूराम माहोर जी के द्वारा बताया गया हे कि यह पे चेकिंग बराबर हो रही है (01:53 Am)", Remark3 = "नाथूराम माहोर जी के द्वारा बताया गया हे कि यह पे चेकिंग बराबर हो रही है (03:53 Am)", Remark4 = "नाथूराम माहोर जी के द्वारा बताया गया हे कि यह पे चेकिंग बराबर हो रही है (05:53 Am)" },
                new { SrNo = 7, Location = "Motijheel", SSTNames = "Kantaprasad Nayak\nR.V.S Narwariya\nRajeev Kumar", Mobile = "9009700695\n7974885212\n9425114402", CheckingTime = "", Remark1 = "vkj-oh-,l ujcjh;k जी के द्वारा बताया गया है की चेकिंग बराबर की जा रही है  (11:55pm)", Remark2 = "vkj-oh-,l ujcjh;k जी के द्वारा बताया गया है की चेकिंग बराबर की जा रही है  (01:55Am)", Remark3 = "vkj-oh-,l ujcjh;k जी के द्वारा बताया गया है की चेकिंग बराबर की जा रही है  (03:55Am)", Remark4 = "vkj-oh-,l ujcjh;k जी के द्वारा फ़ोन नहीं उठाया गया है  (05:55Am)" },
                new { SrNo = 8, Location = "Jalalpur", SSTNames = "Raghvendra Sharma\nSanjay Kumar\nMukesh Kumar Saxena\nR.B sharma\nR.D. rathor\nDinesh sharma", Mobile = "9755524467\n9977083266\n9827620141\n9926228312\n7000356329\n9926957251", CheckingTime = "", Remark1 = "lat; dqekj जी के द्वारा बताया गया है की चेकिंग बराबर हो रही है। (11:57pm)", Remark2 = "lat; dqekj जी के द्वारा बताया गया है की चेकिंग बराबर हो रही है। (01:57Am)", Remark3 = "lat; dqekj जी के द्वारा बताया गया है की चेकिंग बराबर हो रही है। (03:57Am)", Remark4 = "lat; dqekj जी के द्वारा बताया गया है की चेकिंग बराबर हो रही है। (05:57Am)" },
                new { SrNo = 9, Location = "Choudhary dhawa", SSTNames = "Vijayanand Narwariya\nShailendar Mahor\nHemant Mehuriya\nAnil Kumar Pandey", Mobile = "9425630247\n9827366551\n9893836854\n9926232233", CheckingTime = "", Remark1 = "शेलन्‍द्र जी के द्वारा बताया गया हे कि यह पे चेकिंग बराबर हो रही है (11:59 pm)", Remark2 = "शेलन्‍द्र जी के द्वारा बताया गया हे कि यह पे चेकिंग बराबर हो रही है (01:59 Am)", Remark3 = "शेलन्‍द्र जी के द्वारा बताया गया हे कि यह पे चेकिंग बराबर हो रही है (03:59 Am)", Remark4 = "शेलन्‍द्र जी के द्वारा बताया गया हे कि यह पे चेकिंग बराबर हो रही है (05:59 Am)" },
                new { SrNo = 10, Location = "gol pahadiya", SSTNames = "Naresh Verma\nN.P. Singh\nSudhakar Shakya", Mobile = "9926577859\n9425756668\n9981232464", CheckingTime = "", Remark1 = "ujs'k oekZ जी के द्वारा बताया गया है की यह पे चेकिंग बराबर हो रही है (12:02 Am)", Remark2 = "ujs'k oekZ जी के द्वारा बताया गया है की यह पे चेकिंग बराबर हो रही है (02:01 Am)", Remark3 = "ujs'k oekZ जी के द्वारा बताया गया है की यह पे चेकिंग बराबर हो रही है (03:01 Am)", Remark4 = "ujs'k oekZ जी के द्वारा द्वारा फ़ोन नहीं उठाया गया है (06:00 Am)" },
                new { SrNo = 11, Location = "Mohanpur", SSTNames = "Chandrashekhar Pathak\nAPS Yadav\nAnil Verma", Mobile = "8349835329\n9425619969\n8770812076", CheckingTime = "", Remark1 = "अनील वर्मा जी के द्वारा बताया गया है की यह पे चेकिंग बराबर हो रही है (12:05 Am)", Remark2 = "अनील वर्मा जी के द्वारा बताया गया है की यह पे चेकिंग बराबर हो रही है (02:03 Am)", Remark3 = "अनील वर्मा जी के द्वारा बताया गया है की यह पे चेकिंग बराबर हो रही है (03:03 Am)", Remark4 = "अनील वर्मा जी के द्वारा बताया गया है की यह पे चेकिंग बराबर हो रही है (06:04 Am)" },
                new { SrNo = 12, Location = "Bilhati", SSTNames = "MR Dadoriya\nLakhan Singh Raghuvanshi\nAkhilesh Singh Kushwah", Mobile = "9993473072\n9827071354\n8966825212", CheckingTime = "", Remark1 = "अखिलेश कुशवाह जी के द्वारा बताया गया है की चेकिंग बराबर हो रही है।(12:09 Am)", Remark2 = "अखिलेश कुशवाह जी के द्वारा बताया गया है की चेकिंग बराबर हो रही है।(02:05 Am)", Remark3 = "अखिलेश कुशवाह जी के द्वारा बताया गया है की चेकिंग बराबर हो रही है।(03:04 Am)", Remark4 = "अखिलेश कुशवाह जी के द्वारा बताया गया है की चेकिंग बराबर हो रही है।(06:05 Am)" },
                new { SrNo = 13, Location = "kulaith", SSTNames = "Dheerendra Agarwal\nHarnam Singh Patel\nRoopnaryan Sharma", Mobile = "9826547310\n9425709901\n9926593576", CheckingTime = "", Remark1 = "हरनाम सिहं जी के द्वारा बताया गया है कि चेकिंग बराबर हो रही है। (12:10 Am)", Remark2 = "हरनाम सिहं जी के द्वारा बताया गया है कि चेकिंग बराबर हो रही है। (02:09 Am)", Remark3 = "हरनाम सिहं जी के द्वारा बताया गया है कि चेकिंग बराबर हो रही है। (03:08 Am)", Remark4 = "हरनाम सिहं जी के द्वारा बताया गया है कि चेकिंग बराबर हो रही है। (06:08 Am)" },
                new { SrNo = 14, Location = "Sirolaha tiraha", SSTNames = "Ramcharan Kirar\nS.K Bariya\nGajendra Raipuriya", Mobile = "(8878869341/ 8223835689)\n8269225363\n9893896649", CheckingTime = "", Remark1 = "एस के बरैया जी के द्वारा बताया गया है कि चेकिंग बराबर हो रही है। (12:14 Am)", Remark2 = "एस के बरैया जी के द्वारा बताया गया है कि चेकिंग बराबर हो रही है। (02:15 Am)", Remark3 = "एस के बरैया जी के द्वारा बताया गया है कि चेकिंग बराबर हो रही है। (03:14 Am)", Remark4 = "एस के बरैया जी के द्वारा बताया गया है कि चेकिंग बराबर हो रही है। (06:15 Am)" },
                new { SrNo = 15, Location = "Airport tiraha", SSTNames = "Brajesh Singh Bhadoriya\nPrakash Gupta\ndipendra\nSuresh Kumar Bariya\nRajendra Singh", Mobile = "8720895817\n7974733291\n6263824026\n9826229634\n9826506754", CheckingTime = "", Remark1 = "प्रकाश गुप्ता जी के द्वारा बताया गया है की की चेकिंग बराबर हो रही है। (12:15Am)", Remark2 = "प्रकाश गुप्ता जी के द्वारा बताया गया है की की चेकिंग बराबर हो रही है। (02:14Am)", Remark3 = "प्रकाश गुप्ता जी के द्वारा बताया गया है की की चेकिंग बराबर हो रही है। (03:12Am)", Remark4 = "प्रकाश गुप्ता जी के द्वारा बताया गया है की की चेकिंग बराबर हो रही है। (06:11Am)" },
                new { SrNo = 16, Location = "Vicky factory", SSTNames = "Rajeev Pandey\nD.P. Shamra\nVinod Sevriya", Mobile = "9425756558\n9425123219\n9425109213", CheckingTime = "", Remark1 = "डी पी शर्मा जी के द्वारा बताया गया है की चेकिंग बराबर हो रही है। (12:16Am)", Remark2 = "डी पी शर्मा जी के द्वारा बताया गया है की चेकिंग बराबर हो रही है। (02:18Am)", Remark3 = "+", Remark4 = "डी पी शर्मा जी के द्वारा बताया गया है की चेकिंग बराबर हो रही है। (06:12Am)" }
         
                // Add more rows as needed...
            };
            return View(data);
        }

        public IActionResult GeneratePdf()
        {
            // Sample data for the table
            var data = new[]
            {
                new { SrNo = 1, Location = "Mohana", SSTNames = "Rakesh kumar Gupta\nG.K Shrivastav\nOmprakash Kane", Mobile = "9926225166\n9977021658\n9926245820", CheckingTime = "24 hours", Remark1 = "जी के श्रीवास्तव जी के द्वारा बताया गया है की चेकिंग बराबर की जा रही है। (11:42PM)", Remark2 = "जी के श्रीवास्तव जी के द्वारा बताया गया है की चेकिंग बराबर की जा रही है। (01:41AM)", Remark3 = "जी के श्रीवास्तव जी के द्वारा फोर नहीं उठाया गया है। (03:40AM)", Remark4 = "जी के श्रीवास्तव जी के द्वारा बताया गया है की चेकिंग बराबर की जा रही है। (05:42AM)" },
                new { SrNo = 2, Location = "Gohinda", SSTNames = "Piyush Chand\nRK Kokwani\nKapil Kumar", Mobile = "8124484832\n6232641721\n7906394885", CheckingTime = "", Remark1 = "पीयूष चंदजी जी के द्वारा बताया गया है कि चैकिंग बराबर हो रही है।  (11:43 pm)", Remark2 = "पीयूष चंदजी जी के द्वारा बताया गया है कि चैकिंग बराबर हो रही है।  (01:43 Am)", Remark3 = "पीयूष चंदजी जी के द्वारा बताया गया है कि चैकिंग बराबर हो रही है।  (03:42 Am)", Remark4 = "पीयूष चंदजी जी के द्वारा बताया गया है कि चैकिंग बराबर हो रही है।  (05:43 Am)" },
                new { SrNo = 3, Location = "Lodi Tiraha", SSTNames = "Lawrence kumar bodh\nM.C Sharma\nAjay Kumar", Mobile = "9993277488\n9425619912\n9406972494", CheckingTime = "", Remark1 = "एमसी शर्माजी के द्वारा बताया गया है कि चैकिंग बराबर हो रही है।  (11:46 pm)", Remark2 = "एमसी शर्माजी के द्वारा बताया गया है कि चैकिंग बराबर हो रही है।  (01:46 Am)", Remark3 = "ऍम सी शर्मा जी के द्वारा फोर नहीं उठाया गया है।  (03:46 Am)", Remark4 = "ऍम सी शर्मा जी के बताया गया है कि चैकिंग बराबर हो रही है।  (05:46 Am)" },
                new { SrNo = 4, Location = "Ginjora", SSTNames = "Ramesvar singh Alapuriya\nSunil Saxena\nJaynaryan Rajoriya", Mobile = "9926228570\n9826565033\n9926730423", CheckingTime = "", Remark1 = "जयनारायण राजोरिया जी के द्वारा बताया गया है कि चैकिंग बराबर हो रही है।(11:48 pm)", Remark2 = "जयनारायण राजोरिया जी के द्वारा बताया गया है कि चैकिंग बराबर हो रही है।(01:47 Am)", Remark3 = "जयनारायण राजोरिया जी के द्वारा बताया गया है कि चैकिंग बराबर हो रही है।(03:47 Am)", Remark4 = "जयनारायण राजोरिया जी के द्वारा बताया गया है कि चैकिंग बराबर हो रही है।(05:47 Am)" },
                 new { SrNo = 5, Location = "Lidhora", SSTNames = "Shubham Kumar Gupta\nVijay Kumar Sharma\nBhola Gurjar", Mobile = "8120995050\n9753139003\n9630757536", CheckingTime = "", Remark1 = "विजय कुमार शर्मा जी के द्वारा बताया गया हे कि यह पे चेकिंग बराबर हो रही है (11:50pm)", Remark2 = "विजय कुमार शर्मा जी के द्वारा बताया गया हे कि यह पे चेकिंग बराबर हो रही है (01:50Am)", Remark3 = "विजय कुमार शर्मा जी के द्वारा फ़ोन नहीं उठाया गया है (03:50Am)", Remark4 = "विजय कुमार शर्मा जी के द्वारा फ़ोन नहीं उठाया गया है (05:50Am)" },
                new { SrNo = 6, Location = "Chandpura", SSTNames = "Padmakar Mokhribale\nNathuram Mahor\nSatyprakash Samoliya", Mobile = "9826240623\n9826280044\n9893601369", CheckingTime = "", Remark1 = "नाथूराम माहोर जी के द्वारा बताया गया हे कि यह पे चेकिंग बराबर हो रही है (11:53 pm)", Remark2 = "नाथूराम माहोर जी के द्वारा बताया गया हे कि यह पे चेकिंग बराबर हो रही है (01:53 Am)", Remark3 = "नाथूराम माहोर जी के द्वारा बताया गया हे कि यह पे चेकिंग बराबर हो रही है (03:53 Am)", Remark4 = "नाथूराम माहोर जी के द्वारा बताया गया हे कि यह पे चेकिंग बराबर हो रही है (05:53 Am)" },
                new { SrNo = 7, Location = "Motijheel", SSTNames = "Kantaprasad Nayak\nR.V.S Narwariya\nRajeev Kumar", Mobile = "9009700695\n7974885212\n9425114402", CheckingTime = "", Remark1 = "vkj-oh-,l ujcjh;k जी के द्वारा बताया गया है की चेकिंग बराबर की जा रही है  (11:55pm)", Remark2 = "vkj-oh-,l ujcjh;k जी के द्वारा बताया गया है की चेकिंग बराबर की जा रही है  (01:55Am)", Remark3 = "vkj-oh-,l ujcjh;k जी के द्वारा बताया गया है की चेकिंग बराबर की जा रही है  (03:55Am)", Remark4 = "vkj-oh-,l ujcjh;k जी के द्वारा फ़ोन नहीं उठाया गया है  (05:55Am)" },
                new { SrNo = 8, Location = "Jalalpur", SSTNames = "Raghvendra Sharma\nSanjay Kumar\nMukesh Kumar Saxena\nR.B sharma\nR.D. rathor\nDinesh sharma", Mobile = "9755524467\n9977083266\n9827620141\n9926228312\n7000356329\n9926957251", CheckingTime = "", Remark1 = "lat; dqekj जी के द्वारा बताया गया है की चेकिंग बराबर हो रही है। (11:57pm)", Remark2 = "lat; dqekj जी के द्वारा बताया गया है की चेकिंग बराबर हो रही है। (01:57Am)", Remark3 = "lat; dqekj जी के द्वारा बताया गया है की चेकिंग बराबर हो रही है। (03:57Am)", Remark4 = "lat; dqekj जी के द्वारा बताया गया है की चेकिंग बराबर हो रही है। (05:57Am)" },
                new { SrNo = 9, Location = "Choudhary dhawa", SSTNames = "Vijayanand Narwariya\nShailendar Mahor\nHemant Mehuriya\nAnil Kumar Pandey", Mobile = "9425630247\n9827366551\n9893836854\n9926232233", CheckingTime = "", Remark1 = "शेलन्‍द्र जी के द्वारा बताया गया हे कि यह पे चेकिंग बराबर हो रही है (11:59 pm)", Remark2 = "शेलन्‍द्र जी के द्वारा बताया गया हे कि यह पे चेकिंग बराबर हो रही है (01:59 Am)", Remark3 = "शेलन्‍द्र जी के द्वारा बताया गया हे कि यह पे चेकिंग बराबर हो रही है (03:59 Am)", Remark4 = "शेलन्‍द्र जी के द्वारा बताया गया हे कि यह पे चेकिंग बराबर हो रही है (05:59 Am)" },
                new { SrNo = 10, Location = "gol pahadiya", SSTNames = "Naresh Verma\nN.P. Singh\nSudhakar Shakya", Mobile = "9926577859\n9425756668\n9981232464", CheckingTime = "", Remark1 = "ujs'k oekZ जी के द्वारा बताया गया है की यह पे चेकिंग बराबर हो रही है (12:02 Am)", Remark2 = "ujs'k oekZ जी के द्वारा बताया गया है की यह पे चेकिंग बराबर हो रही है (02:01 Am)", Remark3 = "ujs'k oekZ जी के द्वारा बताया गया है की यह पे चेकिंग बराबर हो रही है (03:01 Am)", Remark4 = "ujs'k oekZ जी के द्वारा द्वारा फ़ोन नहीं उठाया गया है (06:00 Am)" },
                new { SrNo = 11, Location = "Mohanpur", SSTNames = "Chandrashekhar Pathak\nAPS Yadav\nAnil Verma", Mobile = "8349835329\n9425619969\n8770812076", CheckingTime = "", Remark1 = "अनील वर्मा जी के द्वारा बताया गया है की यह पे चेकिंग बराबर हो रही है (12:05 Am)", Remark2 = "अनील वर्मा जी के द्वारा बताया गया है की यह पे चेकिंग बराबर हो रही है (02:03 Am)", Remark3 = "अनील वर्मा जी के द्वारा बताया गया है की यह पे चेकिंग बराबर हो रही है (03:03 Am)", Remark4 = "अनील वर्मा जी के द्वारा बताया गया है की यह पे चेकिंग बराबर हो रही है (06:04 Am)" },
                new { SrNo = 12, Location = "Bilhati", SSTNames = "MR Dadoriya\nLakhan Singh Raghuvanshi\nAkhilesh Singh Kushwah", Mobile = "9993473072\n9827071354\n8966825212", CheckingTime = "", Remark1 = "अखिलेश कुशवाह जी के द्वारा बताया गया है की चेकिंग बराबर हो रही है।(12:09 Am)", Remark2 = "अखिलेश कुशवाह जी के द्वारा बताया गया है की चेकिंग बराबर हो रही है।(02:05 Am)", Remark3 = "अखिलेश कुशवाह जी के द्वारा बताया गया है की चेकिंग बराबर हो रही है।(03:04 Am)", Remark4 = "अखिलेश कुशवाह जी के द्वारा बताया गया है की चेकिंग बराबर हो रही है।(06:05 Am)" },
                new { SrNo = 13, Location = "kulaith", SSTNames = "Dheerendra Agarwal\nHarnam Singh Patel\nRoopnaryan Sharma", Mobile = "9826547310\n9425709901\n9926593576", CheckingTime = "", Remark1 = "हरनाम सिहं जी के द्वारा बताया गया है कि चेकिंग बराबर हो रही है। (12:10 Am)", Remark2 = "हरनाम सिहं जी के द्वारा बताया गया है कि चेकिंग बराबर हो रही है। (02:09 Am)", Remark3 = "हरनाम सिहं जी के द्वारा बताया गया है कि चेकिंग बराबर हो रही है। (03:08 Am)", Remark4 = "हरनाम सिहं जी के द्वारा बताया गया है कि चेकिंग बराबर हो रही है। (06:08 Am)" },
                new { SrNo = 14, Location = "Sirolaha tiraha", SSTNames = "Ramcharan Kirar\nS.K Bariya\nGajendra Raipuriya", Mobile = "(8878869341/ 8223835689)\n8269225363\n9893896649", CheckingTime = "", Remark1 = "एस के बरैया जी के द्वारा बताया गया है कि चेकिंग बराबर हो रही है। (12:14 Am)", Remark2 = "एस के बरैया जी के द्वारा बताया गया है कि चेकिंग बराबर हो रही है। (02:15 Am)", Remark3 = "एस के बरैया जी के द्वारा बताया गया है कि चेकिंग बराबर हो रही है। (03:14 Am)", Remark4 = "एस के बरैया जी के द्वारा बताया गया है कि चेकिंग बराबर हो रही है। (06:15 Am)" },
                new { SrNo = 15, Location = "Airport tiraha", SSTNames = "Brajesh Singh Bhadoriya\nPrakash Gupta\ndipendra\nSuresh Kumar Bariya\nRajendra Singh", Mobile = "8720895817\n7974733291\n6263824026\n9826229634\n9826506754", CheckingTime = "", Remark1 = "प्रकाश गुप्ता जी के द्वारा बताया गया है की की चेकिंग बराबर हो रही है। (12:15Am)", Remark2 = "प्रकाश गुप्ता जी के द्वारा बताया गया है की की चेकिंग बराबर हो रही है। (02:14Am)", Remark3 = "प्रकाश गुप्ता जी के द्वारा बताया गया है की की चेकिंग बराबर हो रही है। (03:12Am)", Remark4 = "प्रकाश गुप्ता जी के द्वारा बताया गया है की की चेकिंग बराबर हो रही है। (06:11Am)" },
                new { SrNo = 16, Location = "Vicky factory", SSTNames = "Rajeev Pandey\nD.P. Shamra\nVinod Sevriya", Mobile = "9425756558\n9425123219\n9425109213", CheckingTime = "", Remark1 = "डी पी शर्मा जी के द्वारा बताया गया है की चेकिंग बराबर हो रही है। (12:16Am)", Remark2 = "डी पी शर्मा जी के द्वारा बताया गया है की चेकिंग बराबर हो रही है। (02:18Am)", Remark3 = "+", Remark4 = "डी पी शर्मा जी के द्वारा बताया गया है की चेकिंग बराबर हो रही है। (06:12Am)" }
         
                // Add more rows as needed...
            };

            // Create a memory stream to hold the PDF content
            using (var memoryStream = new MemoryStream())
            {
                // Create a PDF writer and document
                var writer = new PdfWriter(memoryStream);
                var pdf = new PdfDocument(writer);
                var document = new iText.Layout.Document(pdf);

                // Add a title
                document.Add(new Paragraph("SST Data Report")
                    .SetTextAlignment(TextAlignment.CENTER)
                    .SetFontSize(18));

                // Create a table with 8 columns (adjust as per your data)
                var table = new Table(8, true)
                    .SetWidth(100);

                // Add headers for the columns
                table.AddHeaderCell("Sr No.");
                table.AddHeaderCell("SST Location");
                table.AddHeaderCell("SST Name");
                table.AddHeaderCell("Mobile No.");
                table.AddHeaderCell("Checking Time");
                table.AddHeaderCell("Remark & Timing (12:00 AM)");
                table.AddHeaderCell("Remark & Timing (02:00 AM)");
                table.AddHeaderCell("Remark & Timing (04:00 AM)");

                // Add data to the table
                foreach (var row in data)
                {
                    table.AddCell(row.SrNo.ToString());
                    table.AddCell(row.Location);
                    table.AddCell(row.SSTNames);
                    table.AddCell(row.Mobile);
                    table.AddCell(row.CheckingTime);
                    table.AddCell(row.Remark1);
                    table.AddCell(row.Remark2);
                    table.AddCell(row.Remark3);
                }

                // Add the table to the document
                document.Add(table);

                // Finalize the document
                document.Close();

                // Return the PDF as a file result
                return File(memoryStream.ToArray(), "application/pdf", "SSTDataReport.pdf");
            }
        }

        public IActionResult ExportToExcel()
        {
            // Sample data to export (replace with your actual data)
            var data = new List<SSTData>
            {
                new SSTData { SrNo = 1, Location = "Mohana", SSTNames = "Rakesh kumar Gupta\nG.K Shrivastav\nOmprakash Kane", Mobile = "9926225166\n9977021658\n9926245820", CheckingTime = "24 hours", Remark1 = "जी के श्रीवास्तव जी के द्वारा बताया गया है की चेकिंग बराबर की जा रही है। (11:42PM)", Remark2 = "जी के श्रीवास्तव जी के द्वारा बताया गया है की चेकिंग बराबर की जा रही है। (01:41AM)", Remark3 = "जी के श्रीवास्तव जी के द्वारा फोर नहीं उठाया गया है। (03:40AM)", Remark4 = "जी के श्रीवास्तव जी के द्वारा बताया गया है की चेकिंग बराबर की जा रही है। (05:42AM)" },
                new SSTData { SrNo = 2, Location = "Gohinda", SSTNames = "Piyush Chand\nRK Kokwani\nKapil Kumar", Mobile = "8124484832\n6232641721\n7906394885", CheckingTime = "", Remark1 = "पीयूष चंदजी जी के द्वारा बताया गया है कि चैकिंग बराबर हो रही है।  (11:43 pm)", Remark2 = "पीयूष चंदजी जी के द्वारा बताया गया है कि चैकिंग बराबर हो रही है।  (01:43 Am)", Remark3 = "पीयूष चंदजी जी के द्वारा बताया गया है कि चैकिंग बराबर हो रही है।  (03:42 Am)", Remark4 = "पीयूष चंदजी जी के द्वारा बताया गया है कि चैकिंग बराबर हो रही है।  (05:43 Am)" },
                new SSTData { SrNo = 3, Location = "Lodi Tiraha", SSTNames = "Lawrence kumar bodh\nM.C Sharma\nAjay Kumar", Mobile = "9993277488\n9425619912\n9406972494", CheckingTime = "", Remark1 = "एमसी शर्माजी के द्वारा बताया गया है कि चैकिंग बराबर हो रही है।  (11:46 pm)", Remark2 = "एमसी शर्माजी के द्वारा बताया गया है कि चैकिंग बराबर हो रही है।  (01:46 Am)", Remark3 = "ऍम सी शर्मा जी के द्वारा फोर नहीं उठाया गया है।  (03:46 Am)", Remark4 = "ऍम सी शर्मा जी के बताया गया है कि चैकिंग बराबर हो रही है।  (05:46 Am)" },
                new SSTData { SrNo = 4, Location = "Ginjora", SSTNames = "Ramesvar singh Alapuriya\nSunil Saxena\nJaynaryan Rajoriya", Mobile = "9926228570\n9826565033\n9926730423", CheckingTime = "", Remark1 = "जयनारायण राजोरिया जी के द्वारा बताया गया है कि चैकिंग बराबर हो रही है।(11:48 pm)", Remark2 = "जयनारायण राजोरिया जी के द्वारा बताया गया है कि चैकिंग बराबर हो रही है।(01:47 Am)", Remark3 = "जयनारायण राजोरिया जी के द्वारा बताया गया है कि चैकिंग बराबर हो रही है।(03:47 Am)", Remark4 = "जयनारायण राजोरिया जी के द्वारा बताया गया है कि चैकिंग बराबर हो रही है।(05:47 Am)" },
                new SSTData { SrNo = 5, Location = "Lidhora", SSTNames = "Shubham Kumar Gupta\nVijay Kumar Sharma\nBhola Gurjar", Mobile = "8120995050\n9753139003\n9630757536", CheckingTime = "", Remark1 = "विजय कुमार शर्मा जी के द्वारा बताया गया हे कि यह पे चेकिंग बराबर हो रही है (11:50pm)", Remark2 = "विजय कुमार शर्मा जी के द्वारा बताया गया हे कि यह पे चेकिंग बराबर हो रही है (01:50Am)", Remark3 = "विजय कुमार शर्मा जी के द्वारा फ़ोन नहीं उठाया गया है (03:50Am)", Remark4 = "विजय कुमार शर्मा जी के द्वारा फ़ोन नहीं उठाया गया है (05:50Am)" },
                new SSTData { SrNo = 6, Location = "Chandpura", SSTNames = "Padmakar Mokhribale\nNathuram Mahor\nSatyprakash Samoliya", Mobile = "9826240623\n9826280044\n9893601369", CheckingTime = "", Remark1 = "नाथूराम माहोर जी के द्वारा बताया गया हे कि यह पे चेकिंग बराबर हो रही है (11:53 pm)", Remark2 = "नाथूराम माहोर जी के द्वारा बताया गया हे कि यह पे चेकिंग बराबर हो रही है (01:53 Am)", Remark3 = "नाथूराम माहोर जी के द्वारा बताया गया हे कि यह पे चेकिंग बराबर हो रही है (03:53 Am)", Remark4 = "नाथूराम माहोर जी के द्वारा बताया गया हे कि यह पे चेकिंग बराबर हो रही है (05:53 Am)" },
                new SSTData { SrNo = 7, Location = "Motijheel", SSTNames = "Kantaprasad Nayak\nR.V.S Narwariya\nRajeev Kumar", Mobile = "9009700695\n7974885212\n9425114402", CheckingTime = "", Remark1 = "vkj-oh-,l ujcjh;k जी के द्वारा बताया गया है की चेकिंग बराबर की जा रही है  (11:55pm)", Remark2 = "vkj-oh-,l ujcjh;k जी के द्वारा बताया गया है की चेकिंग बराबर की जा रही है  (01:55Am)", Remark3 = "vkj-oh-,l ujcjh;k जी के द्वारा बताया गया है की चेकिंग बराबर की जा रही है  (03:55Am)", Remark4 = "vkj-oh-,l ujcjh;k जी के द्वारा फ़ोन नहीं उठाया गया है  (05:55Am)" },
                new SSTData { SrNo = 8, Location = "Jalalpur", SSTNames = "Raghvendra Sharma\nSanjay Kumar\nMukesh Kumar Saxena\nR.B sharma\nR.D. rathor\nDinesh sharma", Mobile = "9755524467\n9977083266\n9827620141\n9926228312\n7000356329\n9926957251", CheckingTime = "", Remark1 = "lat; dqekj जी के द्वारा बताया गया है की चेकिंग बराबर हो रही है। (11:57pm)", Remark2 = "lat; dqekj जी के द्वारा बताया गया है की चेकिंग बराबर हो रही है। (01:57Am)", Remark3 = "lat; dqekj जी के द्वारा बताया गया है की चेकिंग बराबर हो रही है। (03:57Am)", Remark4 = "lat; dqekj जी के द्वारा बताया गया है की चेकिंग बराबर हो रही है। (05:57Am)" },
                new SSTData { SrNo = 9, Location = "Choudhary dhawa", SSTNames = "Vijayanand Narwariya\nShailendar Mahor\nHemant Mehuriya\nAnil Kumar Pandey", Mobile = "9425630247\n9827366551\n9893836854\n9926232233", CheckingTime = "", Remark1 = "शेलन्‍द्र जी के द्वारा बताया गया हे कि यह पे चेकिंग बराबर हो रही है (11:59 pm)", Remark2 = "शेलन्‍द्र जी के द्वारा बताया गया हे कि यह पे चेकिंग बराबर हो रही है (01:59 Am)", Remark3 = "शेलन्‍द्र जी के द्वारा बताया गया हे कि यह पे चेकिंग बराबर हो रही है (03:59 Am)", Remark4 = "शेलन्‍द्र जी के द्वारा बताया गया हे कि यह पे चेकिंग बराबर हो रही है (05:59 Am)" },
                new SSTData { SrNo = 10, Location = "gol pahadiya", SSTNames = "Naresh Verma\nN.P. Singh\nSudhakar Shakya", Mobile = "9926577859\n9425756668\n9981232464", CheckingTime = "", Remark1 = "ujs'k oekZ जी के द्वारा बताया गया है की यह पे चेकिंग बराबर हो रही है (12:02 Am)", Remark2 = "ujs'k oekZ जी के द्वारा बताया गया है की यह पे चेकिंग बराबर हो रही है (02:01 Am)", Remark3 = "ujs'k oekZ जी के द्वारा बताया गया है की यह पे चेकिंग बराबर हो रही है (03:01 Am)", Remark4 = "ujs'k oekZ जी के द्वारा द्वारा फ़ोन नहीं उठाया गया है (06:00 Am)" },
                new SSTData { SrNo = 11, Location = "Mohanpur", SSTNames = "Chandrashekhar Pathak\nAPS Yadav\nAnil Verma", Mobile = "8349835329\n9425619969\n8770812076", CheckingTime = "", Remark1 = "अनील वर्मा जी के द्वारा बताया गया है की यह पे चेकिंग बराबर हो रही है (12:05 Am)", Remark2 = "अनील वर्मा जी के द्वारा बताया गया है की यह पे चेकिंग बराबर हो रही है (02:03 Am)", Remark3 = "अनील वर्मा जी के द्वारा बताया गया है की यह पे चेकिंग बराबर हो रही है (03:03 Am)", Remark4 = "अनील वर्मा जी के द्वारा बताया गया है की यह पे चेकिंग बराबर हो रही है (06:04 Am)" },
                new SSTData { SrNo = 12, Location = "Bilhati", SSTNames = "MR Dadoriya\nLakhan Singh Raghuvanshi\nAkhilesh Singh Kushwah", Mobile = "9993473072\n9827071354\n8966825212", CheckingTime = "", Remark1 = "अखिलेश कुशवाह जी के द्वारा बताया गया है की चेकिंग बराबर हो रही है।(12:09 Am)", Remark2 = "अखिलेश कुशवाह जी के द्वारा बताया गया है की चेकिंग बराबर हो रही है।(02:05 Am)", Remark3 = "अखिलेश कुशवाह जी के द्वारा बताया गया है की चेकिंग बराबर हो रही है।(03:04 Am)", Remark4 = "अखिलेश कुशवाह जी के द्वारा बताया गया है की चेकिंग बराबर हो रही है।(06:05 Am)" },
                new SSTData { SrNo = 13, Location = "kulaith", SSTNames = "Dheerendra Agarwal\nHarnam Singh Patel\nRoopnaryan Sharma", Mobile = "9826547310\n9425709901\n9926593576", CheckingTime = "", Remark1 = "हरनाम सिहं जी के द्वारा बताया गया है कि चेकिंग बराबर हो रही है। (12:10 Am)", Remark2 = "हरनाम सिहं जी के द्वारा बताया गया है कि चेकिंग बराबर हो रही है। (02:09 Am)", Remark3 = "हरनाम सिहं जी के द्वारा बताया गया है कि चेकिंग बराबर हो रही है। (03:08 Am)", Remark4 = "हरनाम सिहं जी के द्वारा बताया गया है कि चेकिंग बराबर हो रही है। (06:08 Am)" },
                new SSTData { SrNo = 14, Location = "Sirolaha tiraha", SSTNames = "Ramcharan Kirar\nS.K Bariya\nGajendra Raipuriya", Mobile = "(8878869341/ 8223835689)\n8269225363\n9893896649", CheckingTime = "", Remark1 = "एस के बरैया जी के द्वारा बताया गया है कि चेकिंग बराबर हो रही है। (12:14 Am)", Remark2 = "एस के बरैया जी के द्वारा बताया गया है कि चेकिंग बराबर हो रही है। (02:15 Am)", Remark3 = "एस के बरैया जी के द्वारा बताया गया है कि चेकिंग बराबर हो रही है। (03:14 Am)", Remark4 = "एस के बरैया जी के द्वारा बताया गया है कि चेकिंग बराबर हो रही है। (06:15 Am)" },
                new SSTData { SrNo = 15, Location = "Airport tiraha", SSTNames = "Brajesh Singh Bhadoriya\nPrakash Gupta\ndipendra\nSuresh Kumar Bariya\nRajendra Singh", Mobile = "8720895817\n7974733291\n6263824026\n9826229634\n9826506754", CheckingTime = "", Remark1 = "प्रकाश गुप्ता जी के द्वारा बताया गया है की की चेकिंग बराबर हो रही है। (12:15Am)", Remark2 = "प्रकाश गुप्ता जी के द्वारा बताया गया है की की चेकिंग बराबर हो रही है। (02:14Am)", Remark3 = "प्रकाश गुप्ता जी के द्वारा बताया गया है की की चेकिंग बराबर हो रही है। (03:12Am)", Remark4 = "प्रकाश गुप्ता जी के द्वारा बताया गया है की की चेकिंग बराबर हो रही है। (06:11Am)" },
                new SSTData { SrNo = 16, Location = "Vicky factory", SSTNames = "Rajeev Pandey\nD.P. Shamra\nVinod Sevriya", Mobile = "9425756558\n9425123219\n9425109213", CheckingTime = "", Remark1 = "डी पी शर्मा जी के द्वारा बताया गया है की चेकिंग बराबर हो रही है। (12:16Am)", Remark2 = "डी पी शर्मा जी के द्वारा बताया गया है की चेकिंग बराबर हो रही है। (02:18Am)", Remark3 = "+", Remark4 = "डी पी शर्मा जी के द्वारा बताया गया है की चेकिंग बराबर हो रही है। (06:12Am)" }
         
                // Add more data as needed
            };

            // Create an Excel package
            using (var package = new OfficeOpenXml.ExcelPackage())
            {
                // Create the worksheet
                var worksheet = package.Workbook.Worksheets.Add("SST Report");

                // Add headers to the worksheet
                worksheet.Cells[1, 1].Value = "Sr No.";
                worksheet.Cells[1, 2].Value = "SST Location";
                worksheet.Cells[1, 3].Value = "SST Name";
                worksheet.Cells[1, 4].Value = "Mobile No.";
                worksheet.Cells[1, 5].Value = "Checking Time";
                worksheet.Cells[1, 6].Value = "Remark & Timing (12:00 AM)";
                worksheet.Cells[1, 7].Value = "Remark & Timing (02:00 AM)";
                worksheet.Cells[1, 8].Value = "Remark & Timing (03:00 AM)";
                worksheet.Cells[1, 8].Value = "Remark & Timing (06:00 AM)";

                // Populate the worksheet with data
                int row = 2;
                foreach (var item in data)
                {
                    worksheet.Cells[row, 1].Value = item.SrNo;
                    worksheet.Cells[row, 2].Value = item.Location;
                    worksheet.Cells[row, 3].Value = item.SSTNames;
                    worksheet.Cells[row, 4].Value = item.Mobile;
                    worksheet.Cells[row, 5].Value = item.CheckingTime;
                    worksheet.Cells[row, 6].Value = item.Remark1;
                    worksheet.Cells[row, 7].Value = item.Remark2;
                    worksheet.Cells[row, 8].Value = item.Remark3;
                    worksheet.Cells[row, 8].Value = item.Remark4;
                    row++;
                }

                // Format the table (optional)
                worksheet.Cells[1, 1, row - 1, 7].AutoFitColumns();

                // Convert to byte array and return as a file download
                var fileBytes = package.GetAsByteArray();
                return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "SST_Report.xlsx");
            }
        }
        public class SSTData
        {
            public int SrNo { get; set; }
            public string Location { get; set; }
            public string SSTNames { get; set; }
            public string Mobile { get; set; }
            public string CheckingTime { get; set; }
            public string Remark1 { get; set; }
            public string Remark2 { get; set; }
            public string Remark3 { get; set; }
            public string Remark4 { get; set; }
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
