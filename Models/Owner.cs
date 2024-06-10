using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Media.Imaging;

namespace Word_Тепляков.Models
{
    public class Owner
    {
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string SurName { get; set; }
        public int NumberRoom { get; set; }
        public BitmapImage Img { get; set; }
        public bool IsOwner { get; set; }

        public Owner(string FirstName, string LastName, string SurName, int NumberRoom, BitmapImage Img, bool IsOwner)
        { 
            this.FirstName = FirstName;
            this.LastName = LastName;
            this.SurName = SurName;
            this.NumberRoom = NumberRoom;
            this.Img = Img;
            this.IsOwner = IsOwner;
        }
    }
}
