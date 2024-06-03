﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Word_Тепляков.Context;

namespace Word_Тепляков.Elements
{
    /// <summary>
    /// Логика взаимодействия для Owner.xaml
    /// </summary>
    public partial class Owner : UserControl
    {
        public Owner(OwnerContext roomOwner)
        {
            InitializeComponent();
            NameOwner.Content = $"{roomOwner.LastName} {roomOwner.FirstName} {roomOwner.SurName}";
        }
    }
}
