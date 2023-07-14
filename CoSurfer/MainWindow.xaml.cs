using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using OpenQA.Selenium.Edge;
using OpenQA.Selenium.Interactions;
using Microsoft.Win32;
using System.IO;
using System.Threading;
using OfficeOpenXml;
using System.Text;

namespace CoSurfer
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        #region ConsPath
        private readonly string filePath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\Co-SurferSite.txt";
        private readonly string browserInfoPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\Co-SurferBrowser.txt";
        #endregion

        #region Seperator
        readonly char siteSeperator = ',';
        readonly char pathSeperator = ';';
        readonly char writeSeperator = ':';
        readonly char readSeperator = '-';
        #endregion

        #region BoolEdit
        bool writeFromTxtBool = false;
        bool readContentBool = false;
        bool editFromControlBool = false;
        int editSelection = -1;
        #endregion

        #region Var
        List<PathItem> pathList = new List<PathItem>();
        List<Site> siteList = new List<Site>();

        string browserName = "Microsoft Edge";
        string browserLocation;

        IWebDriver driver;
        #endregion

        public MainWindow()
        {
            InitializeComponent();

            ReadFile();

            ReadBrowserInfo();

            RunSitesAutoStart();
        }

        #region Start
        private void ReadFile()
        {
            if (File.Exists(filePath))
            {
                string[] lines = File.ReadAllLines(filePath);

                foreach (var line in lines)
                {
                    string[] siteInfo = line.Split(siteSeperator);
                    Site site = new Site
                    {
                        Title = siteInfo[0],
                        Url = siteInfo[1],
                        AutoWriteLocation = siteInfo[2],
                        SaveLocation = siteInfo[3],
                        AutoStart = siteInfo[4] == "1"
                    };

                    List<PathItem> sitePaths = new List<PathItem>();
                    for (int i = 5; i < siteInfo.Length; i++)
                    {
                        string[] paths = siteInfo[i].Split(pathSeperator);
                        PathItem pathItem = new PathItem
                        {
                            Description = paths[0],
                            Function = paths[1],
                            How = paths[2],
                            Where = paths[3]
                        };
                        sitePaths.Add(pathItem);
                    }
                    site.Paths = sitePaths;

                    siteList.Add(site);

                    
                }
                foreach (var site in siteList)
                {
                    MySitesItem mySiteItem = new MySitesItem
                    {
                        Title = site.Title,
                        Url = site.Url,
                        AutoStart = site.AutoStart
                    };
                    mySitesListBox.Items.Add(mySiteItem);
                }

            }
            else
            {
                GetExample ex = new GetExample();
                string[] lines = ex.example.Split('\n');

                foreach (var line in lines)
                {
                    string[] siteInfo = line.Split(siteSeperator);
                    Site site = new Site
                    {
                        Title = siteInfo[0],
                        Url = siteInfo[1],
                        AutoWriteLocation = siteInfo[2],
                        SaveLocation = siteInfo[3],
                        AutoStart = siteInfo[4] == "1"
                    };

                    List<PathItem> sitePaths = new List<PathItem>();
                    for (int i = 5; i < siteInfo.Length; i++)
                    {
                        string[] paths = siteInfo[i].Split(pathSeperator);
                        PathItem pathItem = new PathItem
                        {
                            Description = paths[0],
                            Function = paths[1],
                            How = paths[2],
                            Where = paths[3]
                        };
                        sitePaths.Add(pathItem);
                    }
                    site.Paths = sitePaths;

                    siteList.Add(site);


                }
                foreach (var site in siteList)
                {
                    MySitesItem mySiteItem = new MySitesItem
                    {
                        Title = site.Title,
                        Url = site.Url,
                        AutoStart = site.AutoStart
                    };
                    mySitesListBox.Items.Add(mySiteItem);
                }
            }
        }
        private void ReadBrowserInfo()
        {
            try
            {
                if (File.Exists(browserInfoPath))
                {
                    string[] lines = File.ReadAllLines(browserInfoPath);

                    browserName = lines[0];
                    browserLocation = lines[1];
                }
            }
            catch (Exception)
            {
                //ignorable
            }
        }
        private void RunSitesAutoStart()
        {
            for (int i = 0; i < siteList.Count; i++)
            {
                if (siteList[i].AutoStart)
                {
                    Site_Click(i);
                }
            }
        }
        #endregion


        #region Add Site And Path Grid
        private void autoWriteTxtFileButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog file = new OpenFileDialog
            {
                Filter = "Text File |*.txt",
                Title = "Select the text file containing the values to be written automatically line by line."
            };
            file.ShowDialog();

            autoWriteLabel.Content = file.FileName;

            //There was file exist control

        }
        private void saveTheReadLocationButton_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog file = new SaveFileDialog
            {
                Filter = "Text File |*.txt",
                Title = "Where do you want to save the read content?"
            };
            file.ShowDialog();

            saveTheReadLabel.Content = file.FileName;
        }
        private void addPathButton_Click(object sender, RoutedEventArgs e)
        {
            pathAddButton.Content = "Add";
            funcAddPathComboBox.SelectedIndex = 0;
            byWhatComboBox.SelectedIndex = 0;
            descriptionTextBox.Text = "Description";
            pathTextBox.Text = "Ex: /html/body/div[1]/div[4]/div/section/div[2]/ul/li[2]/div/span[2]/ins  or //*[@id=\"HTML35\"]\n\nFor \"ReadFor\" or \"ClickFor\" function: /html/body/div[1]/div[4]/div/section/div[2]/ul/li[{ChangeZone}]/div/span[2]/ins or //*[@id=\"HTML{ChangeZone}\"]";
            byWhatWriteTextBox.Text = "write_example";
            writeFromTxtCheckBox.IsChecked = false;

            addPathGrid.Visibility = Visibility.Visible;
        }
        private void delPathButton_Click(object sender, RoutedEventArgs e)
        {
            if (pathListView.SelectedIndex != -1 && pathListView.SelectedItem is PathItem selectedPath)
            {
                pathList.Remove(selectedPath);
                pathListView.ItemsSource = null;
                pathListView.ItemsSource = pathList;
            }
        }
        private void upPathButton_Click(object sender, RoutedEventArgs e)
        {
            if (pathListView.SelectedIndex != -1 && pathListView.SelectedItem is PathItem selectedPath)
            {
                int i = pathList.IndexOf(selectedPath);
                if (i != 0)
                {
                    (pathList[i - 1], pathList[i]) = (pathList[i], pathList[i - 1]);
                    pathListView.ItemsSource = null;
                    pathListView.ItemsSource = pathList;
                }

            }
        }
        private void downPathButton_Click(object sender, RoutedEventArgs e)
        {
            if (pathListView.SelectedIndex != -1 && pathListView.SelectedItem is PathItem selectedPath)
            {
                int i = pathList.IndexOf(selectedPath);
                if (i != pathList.Count - 1)
                {
                    (pathList[i + 1], pathList[i]) = (pathList[i], pathList[i + 1]);
                    pathListView.ItemsSource = null;
                    pathListView.ItemsSource = pathList;
                }


            }
        }
        private void editPathButton_Click(object sender, RoutedEventArgs e)
        {
            if (pathListView.SelectedIndex != -1 && pathListView.SelectedItem is PathItem selectedPath)
            {
                pathAddButton.Content = "Save";
                descriptionTextBox.Text = selectedPath.Description;

                if (selectedPath.Function == "Click")
                    funcAddPathComboBox.SelectedIndex = 0;
                else if (selectedPath.Function.Contains("ReadFor"))
                {
                    funcAddPathComboBox.SelectedIndex = 2;
                    byWhatWriteTextBox.Text = selectedPath.Function.Split(writeSeperator)[1];
                }
                else if (selectedPath.Function.Contains("Read"))
                    funcAddPathComboBox.SelectedIndex = 1;
                else if (selectedPath.Function.Contains("Write"))
                {
                    funcAddPathComboBox.SelectedIndex = 3;

                    if (selectedPath.Function.Split(writeSeperator)[1] == "/FromTxt/")
                    {
                        byWhatWriteTextBox.Text = "";
                        writeFromTxtCheckBox.IsChecked = true;
                    }
                    else
                    {
                        byWhatWriteTextBox.Text = selectedPath.Function.Split(writeSeperator)[1];
                    }
                }
                else if (selectedPath.Function == "Go to url")
                    funcAddPathComboBox.SelectedIndex = 4;
                else if (selectedPath.Function == "Go back")
                    funcAddPathComboBox.SelectedIndex = 5;
                else if (selectedPath.Function == "Go forward")
                    funcAddPathComboBox.SelectedIndex = 6;
                else if (selectedPath.Function == "Wait")
                    funcAddPathComboBox.SelectedIndex = 7;
                else if (selectedPath.Function == "RightClick")
                    funcAddPathComboBox.SelectedIndex = 8;
                else if (selectedPath.Function == "Change Tab")
                    funcAddPathComboBox.SelectedIndex = 9;
                else if (selectedPath.Function == "ClickFor")
                {
                    byWhatWriteTextBox.Text = selectedPath.Function.Split(writeSeperator)[1];
                    funcAddPathComboBox.SelectedIndex = 8;
                }

                if (selectedPath.How == "XPath")
                    byWhatComboBox.SelectedIndex = 0;
                else if (selectedPath.How == "Id")
                    byWhatComboBox.SelectedIndex = 1;
                else if (selectedPath.How == "Name")
                    byWhatComboBox.SelectedIndex = 2;
                else if (selectedPath.How == "ClassName")
                    byWhatComboBox.SelectedIndex = 3;
                else if (selectedPath.How == "TagName")
                    byWhatComboBox.SelectedIndex = 4;


                pathTextBox.Text = selectedPath.Where;

                editSelection = pathListView.SelectedIndex;

                addPathGrid.Visibility = Visibility.Visible;

            }




        }
        private void saveAddSiteButton_Click(object sender, RoutedEventArgs e)
        {
            if (AddSiteValidation())
            {
                Site site = new Site();
                for (int i = 0; i < siteList.Count; i++)
                {
                    if (titleTextBox.Text == siteList[i].Title)
                    {
                        site.Title = titleTextBox.Text + (i + 1).ToString();
                        break;
                    }
                }
                if (site.Title == null)
                {
                    site.Title = titleTextBox.Text;
                }


                site.Url = urlTextBox.Text;
                site.AutoWriteLocation = autoWriteLabel.Content.ToString();
                site.SaveLocation = saveTheReadLabel.Content.ToString();
                site.AutoStart = autoStartCheckbox.IsChecked.Value;
                var temp = new PathItem[pathList.Count];
                pathList.CopyTo(temp);
                site.Paths = temp.ToList<PathItem>();

                siteList.Add(site);
                mySitesListBox.Items.Clear();
                
                foreach (var mySite in siteList)
                {
                    MySitesItem mySiteItem = new MySitesItem
                    {
                        Title = mySite.Title,
                        Url = mySite.Url,
                        AutoStart = mySite.AutoStart
                    };
                    mySitesListBox.Items.Add(mySiteItem);
                }
                

                pathList.Clear();
                addSiteAndPathGrid.Visibility = Visibility.Hidden;
            }
        }
        private bool AddSiteValidation()
        {
            if (titleTextBox.Text?.Length == 0)
            {
                MessageBox.Show("Title is empty!");
                return false;
            }
            else
            {
                foreach (var site in siteList)
                {
                    if (titleTextBox.Text == site.Title)
                    {
                        MessageBox.Show("Warning!\nThere is another site with the same title!");
                        break;
                    }
                }
            }
            if (urlTextBox.Text?.Length == 0)
            {
                MessageBox.Show("Start Url is empty!");
                return false;
            }
            else if (!Uri.IsWellFormedUriString(urlTextBox.Text, UriKind.Absolute))
            {
                MessageBox.Show("Enter valid url!\nExample: https://www.google.com");
                return false;
            }
            if (pathList.Count == 0)
            {
                MessageBox.Show("There is no path!");
                return false;
            }
            else
            {
                foreach (var function in pathList.Select(x => x.Function))
                {
                    if (function == "Write:/FromTxt/")
                    {
                        if (autoWriteLabel.Content.ToString()?.Length == 0 || autoWriteLabel.Content.ToString() == " ")
                        {
                            MessageBox.Show("There is a \"Write:/FromTxt/\" function in paths!\nPlease choose a .txt file.");
                            return false;
                        }
                        else
                        {
                            if (!File.Exists(autoWriteLabel.Content.ToString()))
                            {
                                MessageBox.Show("There is no file on this directory!\n" + autoWriteLabel.Content.ToString());
                                return false;
                            }
                        }
                    }

                    if (function.Contains("Read"))
                    {
                        if (saveTheReadLabel.Content.ToString()?.Length == 0 || saveTheReadLabel.Content.ToString() == " ")
                        {
                            MessageBox.Show("There is a \"Read\" or \"ReadFor\" function in paths!\nPlease choose a .txt file to save.");
                            return false;
                        }
                    }
                }
            }

            return true;
        }
        #endregion


        #region MySites
        private void mySitesListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (mySitesListBox.SelectedIndex != -1)
            {
                editSelection = -1;

                titleTextBox.Text = siteList[mySitesListBox.SelectedIndex].Title;
                urlTextBox.Text = siteList[mySitesListBox.SelectedIndex].Url;
                autoWriteLabel.Content = siteList[mySitesListBox.SelectedIndex].AutoWriteLocation;
                saveTheReadLabel.Content = siteList[mySitesListBox.SelectedIndex].SaveLocation;
                autoStartCheckbox.IsChecked = siteList[mySitesListBox.SelectedIndex].AutoStart;

                pathListView.ItemsSource = null;

                pathListView.ItemsSource = siteList[mySitesListBox.SelectedIndex].Paths;

                var temp = new PathItem[siteList[mySitesListBox.SelectedIndex].Paths.Count];
                siteList[mySitesListBox.SelectedIndex].Paths.CopyTo(temp);
                pathList = temp.ToList<PathItem>();

                addSiteAndPathGrid.Visibility = Visibility.Visible;
            }
        }
        private void mySiteAddButton_Click(object sender, RoutedEventArgs e)
        {
            titleTextBox.Clear();
            urlTextBox.Clear();
            autoWriteLabel.Content = " ";
            saveTheReadLabel.Content = " ";
            pathListView.ItemsSource = null;
            autoStartCheckbox.IsChecked = false;
            pathList.Clear();

            addSiteAndPathGrid.Visibility = Visibility.Visible;
        }
        private void mySiteRemoveButton_Click(object sender, RoutedEventArgs e)
        {
            if (mySitesListBox.SelectedIndex != -1)
            {
                siteList.RemoveAt(mySitesListBox.SelectedIndex);

                mySitesListBox.Items.RemoveAt(mySitesListBox.SelectedIndex);

                pathList.Clear();

                addSiteAndPathGrid.Visibility = Visibility.Hidden;

            }
        }
        private void siteButton_Click(object sender, RoutedEventArgs e)
        {
            var button = (Button)sender;
            var item = (MySitesItem)button.DataContext;

            Site_Click(mySitesListBox.Items.IndexOf(item));
        }
        private void Site_Click(int index)
        {
            try
            {
                if (browserName == "Microsoft Edge")
                {
                    EdgeOptions options = new EdgeOptions();
                    options.PageLoadStrategy = PageLoadStrategy.Eager; //Dont wait for images to load
                    driver = new EdgeDriver(browserLocation, options);
                }
                else if (browserName == "Google Chrome")
                {
                    ChromeOptions options = new ChromeOptions();
                    options.PageLoadStrategy = PageLoadStrategy.Eager; //Dont wait for images to load
                    driver = new ChromeDriver(browserLocation, options);
                }

                if (driver != null)
                {
                    driver.Url = siteList[index].Url;

                    driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(5);
                    driver.Manage().Window.Maximize();

                    if (RunFunction(siteList[index].Paths, driver) == 1)
                    {
                        driver.Quit();
                    }

                }
            }
            catch (Exception)
            {
                MessageBox.Show("There is a problem with the browser or url.");
            }
        }
        #endregion


        #region Window
        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (true)//TODO check that the data to be written to the text file is regular and there is no problem 
            {
                if (File.Exists(filePath))
                {
                    File.Delete(filePath);
                }

                using (StreamWriter file = File.CreateText(filePath))
                {
                    foreach (Site site in siteList)
                    {
                        StringBuilder line = new StringBuilder();
                        line.Append(site.Title);
                        line.Append(siteSeperator);
                        line.Append(site.Url);
                        line.Append(siteSeperator);
                        line.Append(site.AutoWriteLocation);
                        line.Append(siteSeperator);
                        line.Append(site.SaveLocation);
                        line.Append(siteSeperator);
                        line.Append(site.AutoStart ? "1" : "0");
                                                
                        foreach (var path in site.Paths)
                        {
                            line.Append(siteSeperator);
                            line.Append(path.Description);
                            line.Append(pathSeperator);
                            line.Append(path.Function);
                            line.Append(pathSeperator);
                            line.Append(path.How);
                            line.Append(pathSeperator);
                            line.Append(path.Where);
                        }
                        file.WriteLine(line.ToString());
                    }

                    file.Close();
                }

                if (File.Exists(browserInfoPath))
                {
                    File.Delete(browserInfoPath);
                }

                using (StreamWriter file = File.CreateText(browserInfoPath))
                {
                    string line = browserName + "\n" + browserLocation;

                    file.WriteLine(line);

                    file.Close();
                }
            }
            else
            {
                e.Cancel = true;
            }
            
        }
        #endregion


        #region PathAdd
        private void pathAddButton_Click(object sender, RoutedEventArgs e)
        {
            if (PathAddValidation())
            {
                PathItem path = new PathItem();

                path.Description = descriptionTextBox.Text;
                path.Function = funcAddPathComboBox.Text;
                if (path.Function == "Write")
                {
                    path.Function += writeFromTxtCheckBox.IsChecked.GetValueOrDefault(false) ? ":/FromTxt/" : (":" + byWhatWriteTextBox.Text);
                }
                else if (path.Function == "ReadFor" || path.Function == "ClickFor")
                {
                    path.Function += ":" + byWhatWriteTextBox.Text;
                }
                path.How = byWhatComboBox.Text;
                path.Where = pathTextBox.Text;


                

                


                if (editFromControlBool)
                {
                    //editFromControlBool = false;
                    int forIndex;
                    if (path.Function.Contains("ReadFor") || path.Function.Contains("ClickFor"))
                    {
                        string readForText = path.Function.Split(':')[1];
                        forIndex = Convert.ToInt32(readForText.Split(readSeperator)[0]);
                        
                    }
                    else
                    {
                        forIndex = 0;
                    }
                    
                    if (ControlNoSuchElement(path, editSelection, driver, forIndex))
                    {
                        MessageBox.Show("Element found!\nPath edited.\nPlease press Control again.");

                        pathList[editSelection] = path;
                        editSelection = -1;

                        driver.Quit();

                        pathListView.ItemsSource = null;

                        pathListView.ItemsSource = pathList;

                        addPathGrid.Visibility = Visibility.Hidden;
                    }
                }
                else
                {
                    if (editSelection != -1)
                    {
                        pathList[editSelection] = path;
                        editSelection = -1;
                    }
                    else
                    {
                        pathList.Add(path);
                    }

                    pathListView.ItemsSource = null;

                    pathListView.ItemsSource = pathList;

                    addPathGrid.Visibility = Visibility.Hidden;
                }

                

            }
                


        }
        private bool PathAddValidation()
        {
            if (descriptionTextBox.Text?.Length == 0)
            {
                MessageBox.Show("Description is empty!");
                return false;
            }
            if (funcAddPathComboBox.Text != "Go back" && funcAddPathComboBox.Text != "Go forward" && pathTextBox.Text?.Length == 0)
            {
                MessageBox.Show("Path is empty!");
                return false;
            }
            if (funcAddPathComboBox.Text == "ReadFor" || funcAddPathComboBox.Text == "ClickFor")
            {
                if (byWhatWriteTextBox.Text?.Length == 0)
                {
                    MessageBox.Show("Please write start and end index! \r\nExample: 1-13");
                    return false;
                }
                else if (!byWhatWriteTextBox.Text.Contains(readSeperator))
                {
                    MessageBox.Show("Can't find \'" + readSeperator +"\' character!");
                    return false;
                }
                else
                {
                    string[] temp = byWhatWriteTextBox.Text.Split(readSeperator);
                    int start, end;

                    if (temp.Length > 2)
                    {
                        MessageBox.Show("There has to be only one \'" + readSeperator + "\' character!");
                        return false;
                    }
                    else if (!int.TryParse(temp[0], out start) || !int.TryParse(temp[1], out end))
                    {
                        MessageBox.Show("Please enter in the form of number" + readSeperator + "number!");
                        return false;
                    }
                    else if (end <= start)
                    {
                        MessageBox.Show("The number on the left must be lower than on the right!");
                        return false;
                    }
                }

                

                if (!pathTextBox.Text.Contains("{ChangeZone}"))
                {
                    MessageBox.Show("{ChangeZone} must be present in the path!");
                    return false;
                }
            
                
                
            }
            if (funcAddPathComboBox.Text == "Write" && writeFromTxtCheckBox.IsChecked.Value == false && byWhatWriteTextBox.Text?.Length == 0)
            {
                MessageBox.Show("Please enter write content!");
                return false;
            }
            if (funcAddPathComboBox.Text == "Change Tab")
            {
                int s;
                if (!int.TryParse(pathTextBox.Text, out s))
                {
                    MessageBox.Show("Please enter only number!");
                    return false;
                }
            }
            if (funcAddPathComboBox.Text == "Wait")
            {
                int s;
                if (!int.TryParse(pathTextBox.Text, out s))
                {
                    MessageBox.Show("Please enter only number! (miliseconds)");
                    return false;
                }
            }
            if (funcAddPathComboBox.Text == "Go to url")
            {
                if (!Uri.IsWellFormedUriString(pathTextBox.Text, UriKind.Absolute))
                {
                    MessageBox.Show("Enter valid url!\nExample: https://www.google.com");
                    return false;
                }
            }

            if (byWhatComboBox.Visibility == Visibility.Visible)
            {
                if (byWhatComboBox.SelectedValue.ToString() == "XPath")
                {
                    if (pathTextBox.Text[0] != '/')
                    {
                        MessageBox.Show("Enter valid XPath!\n");
                        return false;
                    }
                }
            }

            return true;
        }
        private void funcAddPathComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

            if (pathTextBox != null)
            {
                if (funcAddPathComboBox.SelectedValue.ToString() == "Write")
                {
                    byWhatComboBox.SelectedIndex = 0;
                    pathTextBox.Text = "Ex: /html/body/div[1]/div[4]/div/section/div[2]/ul/li[2]/div/span[2]/ins  or //*[@id=\"HTML35\"]\n\nFor \"ReadFor\" or \"ClickFor\" function: /html/body/div[1]/div[4]/div/section/div[2]/ul/li[{ChangeZone}]/div/span[2]/ins or //*[@id=\"HTML{ChangeZone}\"]";
                    byWhatWriteTextBox.Visibility = Visibility.Visible;
                    writeFromTxtCheckBox.Visibility = Visibility.Visible;
                }
                else if (funcAddPathComboBox.SelectedValue.ToString() == "Go to url")
                {
                    pathTextBox.Text = "Ex: https://www.google.com";
                    byWhatWriteTextBox.Visibility = Visibility.Hidden;
                    writeFromTxtCheckBox.Visibility = Visibility.Hidden;
                    pathTextBox.Visibility = Visibility.Visible;
                    byWhatComboBox.Visibility = Visibility.Hidden;
                }
                else if (funcAddPathComboBox.SelectedValue.ToString() == "Go back")
                {
                    byWhatWriteTextBox.Visibility = Visibility.Hidden;
                    writeFromTxtCheckBox.Visibility = Visibility.Hidden;
                    pathTextBox.Visibility = Visibility.Hidden;
                    byWhatComboBox.Visibility = Visibility.Hidden;
                }
                else if (funcAddPathComboBox.SelectedValue.ToString() == "Go forward")
                {
                    byWhatWriteTextBox.Visibility = Visibility.Hidden;
                    writeFromTxtCheckBox.Visibility = Visibility.Hidden;
                    pathTextBox.Visibility = Visibility.Hidden;
                    byWhatComboBox.Visibility = Visibility.Hidden;
                }
                else if (funcAddPathComboBox.SelectedValue.ToString() == "ReadFor")
                {
                    byWhatComboBox.SelectedIndex = 0;
                    pathTextBox.Text = "Ex: /html/body/div[1]/div[4]/div/section/div[2]/ul/li[2]/div/span[2]/ins  or //*[@id=\"HTML35\"]\n\nFor \"ReadFor\" or \"ClickFor\" function: /html/body/div[1]/div[4]/div/section/div[2]/ul/li[{ChangeZone}]/div/span[2]/ins or //*[@id=\"HTML{ChangeZone}\"]";
                    byWhatComboBox.Visibility = Visibility.Visible;
                    byWhatWriteTextBox.Text = "Ex: 1-13";
                    byWhatWriteTextBox.Visibility = Visibility.Visible;
                    pathTextBox.Visibility = Visibility.Visible;
                }
                else if (funcAddPathComboBox.SelectedValue.ToString() == "RightClick")
                {
                    byWhatComboBox.SelectedIndex = 0;
                    pathTextBox.Text = "Ex: /html/body/div[1]/div[4]/div/section/div[2]/ul/li[2]/div/span[2]/ins  or //*[@id=\"HTML35\"]\n\nFor \"ReadFor\" or \"ClickFor\" function: /html/body/div[1]/div[4]/div/section/div[2]/ul/li[{ChangeZone}]/div/span[2]/ins or //*[@id=\"HTML{ChangeZone}\"]";
                    byWhatComboBox.Visibility = Visibility.Visible;
                    byWhatWriteTextBox.Visibility = Visibility.Hidden;
                    writeFromTxtCheckBox.Visibility = Visibility.Hidden;
                }
                else if (funcAddPathComboBox.SelectedValue.ToString() == "Change Tab")
                {
                    byWhatWriteTextBox.Visibility = Visibility.Hidden;
                    writeFromTxtCheckBox.Visibility = Visibility.Hidden;
                    byWhatComboBox.Visibility = Visibility.Hidden;
                    pathTextBox.Text = "For second tab you need to write \"2\"";
                }
                else if (funcAddPathComboBox.SelectedValue.ToString() == "ClickFor")
                {
                    byWhatComboBox.SelectedIndex = 0;
                    pathTextBox.Text = "Ex: /html/body/div[1]/div[4]/div/section/div[2]/ul/li[2]/div/span[2]/ins  or //*[@id=\"HTML35\"]\n\nFor \"ReadFor\" or \"ClickFor\" function: /html/body/div[1]/div[4]/div/section/div[2]/ul/li[{ChangeZone}]/div/span[2]/ins or //*[@id=\"HTML{ChangeZone}\"]";
                    writeFromTxtCheckBox.Visibility = Visibility.Hidden;
                    byWhatComboBox.Visibility = Visibility.Visible;
                    byWhatWriteTextBox.Text = "Ex: 1-13";
                    byWhatWriteTextBox.Visibility = Visibility.Visible;
                    pathTextBox.Visibility = Visibility.Visible;
                }
                else if (funcAddPathComboBox.SelectedValue.ToString() == "Wait")
                {
                    byWhatWriteTextBox.Visibility = Visibility.Hidden;
                    writeFromTxtCheckBox.Visibility = Visibility.Hidden;
                    pathTextBox.Text = "Ex: 1000 (ms) or -1 for wait until you confirm.";
                    byWhatComboBox.Visibility = Visibility.Hidden;
                    pathTextBox.Visibility = Visibility.Visible;
                }
                else
                {
                    byWhatComboBox.SelectedIndex = 0;
                    pathTextBox.Text = "Ex: /html/body/div[1]/div[4]/div/section/div[2]/ul/li[2]/div/span[2]/ins  or //*[@id=\"HTML35\"]\n\nFor \"ReadFor\" or \"ClickFor\" function: /html/body/div[1]/div[4]/div/section/div[2]/ul/li[{ChangeZone}]/div/span[2]/ins or //*[@id=\"HTML{ChangeZone}\"]";
                    byWhatWriteTextBox.Visibility = Visibility.Hidden;
                    writeFromTxtCheckBox.Visibility = Visibility.Hidden;
                    byWhatComboBox.Visibility = Visibility.Visible;
                    pathTextBox.Visibility = Visibility.Visible;
                }
            }


        }
        private void byWhatComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (pathTextBox != null)
            {
                if (byWhatComboBox.SelectedValue.ToString() == "XPath")
                {
                    pathTextBox.Text = "Ex: /html/body/div[1]/div[4]/div/section/div[2]/ul/li[2]/div/span[2]/ins  or //*[@id=\"HTML35\"]\n\nFor \"ReadFor\" or \"ClickFor\" function: /html/body/div[1]/div[4]/div/section/div[2]/ul/li[{ChangeZone}]/div/span[2]/ins or //*[@id=\"HTML{ChangeZone}\"]";
                }
                if (byWhatComboBox.SelectedValue.ToString() == "Id")
                {
                    pathTextBox.Text = "Ex: NAVBAR_ID_3\n\nFor \"ReadFor\" or \"ClickFor\" function: NAVBAR_ID_{ChangeZone}";
                }
                if (byWhatComboBox.SelectedValue.ToString() == "Name")
                {
                    pathTextBox.Text = "Ex: NAVBAR_NAME_3\n\nFor \"ReadFor\" or \"ClickFor\" function: NAVBAR_NAME_{ChangeZone}";
                }
                if (byWhatComboBox.SelectedValue.ToString() == "ClassName")
                {
                    pathTextBox.Text = "Ex: NAVBAR_CLASS_3\n\nFor \"ReadFor\" or \"ClickFor\" function: NAVBAR_CLASS_{ChangeZone}";
                }
                if (byWhatComboBox.SelectedValue.ToString() == "TagName")
                {
                    pathTextBox.Text = "Ex: NAVBAR_TAG_3\n\nFor \"ReadFor\" or \"ClickFor\" function: NAVBAR_TAG_{ChangeZone}";
                }
            }
        }
        private void pathAddCancelButton_Click(object sender, RoutedEventArgs e)
        {
            if (editFromControlBool)
            {
                editFromControlBool = false;
                editSelection = -1;

                if (driver != null)
                {
                    driver.Quit();
                }
            }
            addPathGrid.Visibility = Visibility.Hidden;
        }
        #endregion


        #region Control
        private void controlAddSiteButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (browserName == "Microsoft Edge")
                {
                    EdgeOptions options = new EdgeOptions();
                    options.PageLoadStrategy = PageLoadStrategy.Eager; //Dont wait for images to load
                    driver = new EdgeDriver(browserLocation, options);
                }
                else if (browserName == "Google Chrome")
                {
                    ChromeOptions options = new ChromeOptions();
                    options.PageLoadStrategy = PageLoadStrategy.Eager; //Dont wait for images to load
                    driver = new ChromeDriver(browserLocation, options);
                }

                if (driver != null)
                {
                    driver.Url = urlTextBox.Text;

                    driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(5);
                    driver.Manage().Window.Maximize();

                    if (RunFunction(pathList, driver) == 1)
                    {
                        driver.Quit();
                    }

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("There is a problem with the browser or url.");
            }


        }
        private int RunFunction(List<PathItem> paths, IWebDriver driver)
        {
            try
            {
                string readContent = "";
                readContentBool = false;
                writeFromTxtBool = false;

                List<string> searchList = new List<string>();

                foreach (var path in paths)
                {
                    if(path.Function == "Write:/FromTxt/")
                    {
                        writeFromTxtBool = true;

                        if (File.Exists(autoWriteLabel.Content.ToString()))
                        {
                            searchList = File.ReadAllLines(autoWriteLabel.Content.ToString()).ToList();
                        }
                        break;
                    }
                    
                    
                }

                if (searchList.Count == 0)
                    searchList.Add("");


                for (int i = 0; i < searchList.Count; i++)
                {
                    for (int j = 0; j < paths.Count; j++)
                    {
                        if (paths[j].Function == "Click")
                        {
                            if (ControlNoSuchElement(paths[j], j, driver))
                            {
                                if (paths[j].How == "XPath")
                                    driver.FindElement(By.XPath(paths[j].Where)).Click();
                                else if (paths[j].How == "Id")
                                    driver.FindElement(By.Id(paths[j].Where)).Click();
                                else if (paths[j].How == "Name")
                                    driver.FindElement(By.Name(paths[j].Where)).Click();
                                else if (paths[j].How == "ClassName")
                                    driver.FindElement(By.ClassName(paths[j].Where)).Click();
                                else if (paths[j].How == "TagName")
                                    driver.FindElement(By.TagName(paths[j].Where)).Click();
                            }
                            else
                                return 0;
                        }
                        else if (paths[j].Function == "Read")
                        {
                            if (ControlNoSuchElement(paths[j], j, driver))
                            {
                                if (paths[j].How == "XPath")
                                    readContent += driver.FindElement(By.XPath(paths[j].Where)).Text + "\n";
                                else if (paths[j].How == "Id")
                                    readContent += driver.FindElement(By.Id(paths[j].Where)).Text + "\n";
                                else if (paths[j].How == "Name")
                                    readContent += driver.FindElement(By.Name(paths[j].Where)).Text + "\n";
                                else if (paths[j].How == "ClassName")
                                    readContent += driver.FindElement(By.ClassName(paths[j].Where)).Text + "\n";
                                else if (paths[j].How == "TagName")
                                    readContent += driver.FindElement(By.TagName(paths[j].Where)).Text + "\n";

                                readContentBool = true;
                            }
                            else
                                return 0;
                        }
                        else if (paths[j].Function.Contains("ReadFor"))
                        {
                            string readForText = paths[j].Function.Split(':')[1];
                            int readForEnd = Convert.ToInt32(readForText.Split(readSeperator)[1]);

                            for (int readForStart = Convert.ToInt32(readForText.Split(readSeperator)[0]); readForStart <= readForEnd; readForStart++)
                            {
                                string tempPath = paths[j].Where;

                                tempPath = tempPath.Replace("{ChangeZone}", readForStart.ToString());

                                if (ControlNoSuchElement(paths[j], j, driver, readForStart))
                                {
                                    if (paths[j].How == "XPath")
                                        readContent += driver.FindElement(By.XPath(tempPath)).Text + "\n";
                                    else if (paths[j].How == "Id")
                                        readContent += driver.FindElement(By.Id(tempPath)).Text + "\n";
                                    else if (paths[j].How == "Name")
                                        readContent += driver.FindElement(By.Name(tempPath)).Text + "\n";
                                    else if (paths[j].How == "ClassName")
                                        readContent += driver.FindElement(By.ClassName(tempPath)).Text + "\n";
                                    else if (paths[j].How == "TagName")
                                        readContent += driver.FindElement(By.TagName(tempPath)).Text + "\n";
                                }
                                else
                                    return 0;
                            }
                            

                            readContentBool = true;
                            
                        }
                        else if (paths[j].Function.Contains("Write"))
                        {
                            string searchText;
                            if (writeFromTxtBool)
                            {
                                searchText = searchList[i];
                            }
                            else
                            {
                                searchText = paths[j].Function.Split(writeSeperator)[1];
                            }

                            if (ControlNoSuchElement(paths[j], j, driver))
                            {
                                if (paths[j].How == "XPath")
                                    driver.FindElement(By.XPath(paths[j].Where)).SendKeys(searchText);
                                else if (paths[j].How == "Id")
                                    driver.FindElement(By.Id(paths[j].Where)).SendKeys(searchText);
                                else if (paths[j].How == "Name")
                                    driver.FindElement(By.Name(paths[j].Where)).SendKeys(searchText);
                                else if (paths[j].How == "ClassName")
                                    driver.FindElement(By.ClassName(paths[j].Where)).SendKeys(searchText);
                                else if (paths[j].How == "TagName")
                                    driver.FindElement(By.TagName(paths[j].Where)).SendKeys(searchText);
                            }
                            else
                                return 0;
                        }
                        else if (paths[j].Function == "Go to url")
                        {
                            driver.Navigate().GoToUrl(paths[j].Where);
                        }
                        else if (paths[j].Function == "Go back")
                        {
                            driver.Navigate().Back();
                        }
                        else if (paths[j].Function == "Go forward")
                        {
                            driver.Navigate().Forward();
                        }
                        else if (paths[j].Function == "Wait")
                        {
                            if (Convert.ToInt32(paths[j].Where) == -1)
                            {
                                this.Topmost = true;
                                this.Topmost = false;
                                MessageBox.Show(this, "Waiting for your command!");
                            }
                            else
                                Thread.Sleep(Convert.ToInt32(paths[j].Where));
                        }
                        else if (paths[j].Function == "RightClick")
                        {
                            Actions actions = new Actions(driver);
                            WebElement elementLocator = null;

                            if (ControlNoSuchElement(paths[j], j, driver))
                            {
                                if (paths[j].How == "XPath")
                                    elementLocator = (WebElement)driver.FindElement(By.XPath(paths[j].Where));
                                else if (paths[j].How == "Id")
                                    elementLocator = (WebElement)driver.FindElement(By.Id(paths[j].Where));
                                else if (paths[j].How == "Name")
                                    elementLocator = (WebElement)driver.FindElement(By.Name(paths[j].Where));
                                else if (paths[j].How == "ClassName")
                                    elementLocator = (WebElement)driver.FindElement(By.ClassName(paths[j].Where));
                                else if (paths[j].How == "TagName")
                                    elementLocator = (WebElement)driver.FindElement(By.TagName(paths[j].Where));

                                if (elementLocator != null)
                                {
                                    actions.ContextClick(elementLocator).Perform();
                                }
                            }
                            else
                                return 0;
                        }
                        else if (paths[j].Function == "Change Tab")
                        {
                            if (driver.WindowHandles.Count > Convert.ToInt32(paths[j].Where) - 1)
                            {
                                driver.SwitchTo().Window(driver.WindowHandles[Convert.ToInt32(paths[j].Where) - 1]);
                            }
                            else
                            {
                                this.Topmost = true;
                                this.Topmost = false;
                                MessageBox.Show(this, "There was no tab!");
                                return 0;
                            }
                        }
                        else if (paths[j].Function.Contains("ClickFor"))
                        {
                            string readForText = paths[j].Function.Split(':')[1];
                            int readForEnd = Convert.ToInt32(readForText.Split(readSeperator)[1]);

                            for (int readForStart = Convert.ToInt32(readForText.Split(readSeperator)[0]); readForStart <= readForEnd; readForStart++)
                            {
                                string tempPath = paths[j].Where;

                                tempPath = tempPath.Replace("{ChangeZone}", readForStart.ToString());

                                if (ControlNoSuchElement(paths[j], j, driver, readForStart))
                                {
                                    if (paths[j].How == "XPath")
                                        driver.FindElement(By.XPath(tempPath)).Click();
                                    else if (paths[j].How == "Id")
                                        driver.FindElement(By.Id(tempPath)).Click();
                                    else if (paths[j].How == "Name")
                                        driver.FindElement(By.Name(tempPath)).Click();
                                    else if (paths[j].How == "ClassName")
                                        driver.FindElement(By.ClassName(tempPath)).Click();
                                    else if (paths[j].How == "TagName")
                                        driver.FindElement(By.TagName(tempPath)).Click();
                                }
                                else
                                    return 0;
                            }

                        }
                    }
                    
                }


                if (readContentBool)
                {
                    if (File.Exists(saveTheReadLabel.Content.ToString()))
                    {
                        File.Delete(saveTheReadLabel.Content.ToString());
                    }

                    using (StreamWriter file = File.CreateText(saveTheReadLabel.Content.ToString()))
                    {
                        file.WriteLine(readContent);

                        file.Close();
                    }
                }


                return 1;
            }
            catch (Exception )
            {
                this.Topmost = true;
                this.Topmost = false;
                MessageBox.Show(this, "Function denied.");
                return 0;
            }

        }
        private bool ControlNoSuchElement(PathItem path, int index, IWebDriver driver, int forIndex = 0)
        {
            List<IWebElement> elementList = new List<IWebElement>();

            string tempPath = path.Where;

            if (path.Function.Contains("ReadFor") || path.Function.Contains("ClickFor"))
            {
                tempPath = tempPath.Replace("{ChangeZone}", forIndex.ToString());
            }


            if (path.How == "XPath")
                elementList.AddRange(driver.FindElements(By.XPath(tempPath)));
            else if (path.How == "Id")
                elementList.AddRange(driver.FindElements(By.Id(tempPath)));
            else if (path.How == "Name")
                elementList.AddRange(driver.FindElements(By.Name(tempPath)));
            else if (path.How == "ClassName")
                elementList.AddRange(driver.FindElements(By.ClassName(tempPath)));
            else if (path.How == "TagName")
                elementList.AddRange(driver.FindElements(By.TagName(tempPath)));

            
            if (elementList.Count > 0)
            {
                editFromControlBool = false;
                return true;
            }
            else
            {
                editFromControlBool = true;

                this.Topmost = true;
                this.Topmost = false;
                MessageBox.Show(this, "Please rewrite element address.\n" + tempPath + "\nCan't access right now.\nOr cancel control.\n\nNote: Since the site takes a long time to fully load, adding a \"Wait\" function may solve the problem.");

                pathAddButton.Content = "Save";

                descriptionTextBox.Text = path.Description;

                if (path.Function == "Click")
                    funcAddPathComboBox.SelectedIndex = 0;
                else if (path.Function.Contains("ReadFor"))
                {
                    funcAddPathComboBox.SelectedIndex = 2;
                    byWhatWriteTextBox.Text = path.Function.Split(writeSeperator)[1];
                }
                else if (path.Function.Contains("Read"))
                    funcAddPathComboBox.SelectedIndex = 1;
                else if (path.Function.Contains("Write"))
                {
                    funcAddPathComboBox.SelectedIndex = 3;

                    if (path.Function.Split(writeSeperator)[1] == "/FromTxt/")
                    {
                        byWhatWriteTextBox.Text = "";
                        writeFromTxtCheckBox.IsChecked = true;
                    }
                    else
                    {
                        byWhatWriteTextBox.Text = path.Function.Split(writeSeperator)[1];
                    }
                }
                else if (path.Function == "Go to url")
                    funcAddPathComboBox.SelectedIndex = 4;
                else if (path.Function == "Go back")
                    funcAddPathComboBox.SelectedIndex = 5;
                else if (path.Function == "Go forward")
                    funcAddPathComboBox.SelectedIndex = 6;
                else if (path.Function == "Wait")
                    funcAddPathComboBox.SelectedIndex = 7;
                else if (path.Function == "RightClick")
                    funcAddPathComboBox.SelectedIndex = 8;
                else if (path.Function == "Change Tab")
                    funcAddPathComboBox.SelectedIndex = 9;
                else if (path.Function == "ClickFor")
                {
                    byWhatWriteTextBox.Text = path.Function.Split(writeSeperator)[1];
                    funcAddPathComboBox.SelectedIndex = 8;
                }

                if (path.How == "XPath")
                    byWhatComboBox.SelectedIndex = 0;
                else if (path.How == "Id")
                    byWhatComboBox.SelectedIndex = 1;
                else if (path.How == "Name")
                    byWhatComboBox.SelectedIndex = 2;
                else if (path.How == "ClassName")
                    byWhatComboBox.SelectedIndex = 3;
                else if (path.How == "TagName")
                    byWhatComboBox.SelectedIndex = 4;


                pathTextBox.Text = path.Where;

                editSelection = index;

                addPathGrid.Visibility = Visibility.Visible;

                this.Topmost = false;

                return false;
            }

        }
        #endregion


        #region WriteFromTxtCheckBox
        private void writeFromTxtCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            byWhatWriteTextBox.IsReadOnly = true;
            byWhatWriteTextBox.IsEnabled = false;
        }
        private void writeFromTxtCheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            byWhatWriteTextBox.IsReadOnly = false;
            byWhatWriteTextBox.IsEnabled = true;
        }
        #endregion


        #region Export Import Excel
        private void exportOneButton_Click(object sender, RoutedEventArgs e)
        {
            if (mySitesListBox.SelectedIndex != -1)
            {
                
                SaveFileDialog file = new SaveFileDialog
                {
                    Filter = "Excel File |*.xlsx",
                    Title = "Where do you want to save the site?"
                };
                file.ShowDialog();
                Site site = siteList[mySitesListBox.SelectedIndex];

                if (file.FileName?.Length != 0)
                {
                    using (var package = new ExcelPackage())
                    {
                        ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Sheet1");

                        worksheet.Cells[1, 1].Value = "Title";
                        worksheet.Cells[1, 2].Value = "Url";
                        worksheet.Cells[1, 3].Value = "AutoWriteLocation";
                        worksheet.Cells[1, 4].Value = "SaveReadLocation";
                        worksheet.Cells[1, 5].Value = "AutoStart";
                        worksheet.Cells[1, 6].Value = "Paths-->";

                        worksheet.Cells[2, 1].Value = site.Title;
                        worksheet.Cells[2, 2].Value = site.Url;
                        worksheet.Cells[2, 3].Value = site.AutoWriteLocation;
                        worksheet.Cells[2, 4].Value = site.SaveLocation;
                        worksheet.Cells[2, 5].Value = site.AutoStart ? "1" : "0";
                        for (int i = 0; i < site.Paths.Count; i++)
                        {
                            worksheet.Cells[2, i + 6].Value = site.Paths[i].Description + pathSeperator +
                                                              site.Paths[i].Function + pathSeperator +
                                                              site.Paths[i].How + pathSeperator +
                                                              site.Paths[i].Where;
                        }
                        FileInfo excelFile = new FileInfo(file.FileName);
                        package.SaveAs(excelFile);

                        MessageBox.Show("Excel file saved in this location:\n" + file.FileName);
                    }

                }
                
            }
            else
            {
                MessageBox.Show("You must choose a site to export!");
            }
        }
        private void exportAllButton_Click(object sender, RoutedEventArgs e)
        {
            if (siteList.Count != 0)
            {

                SaveFileDialog file = new SaveFileDialog
                {
                    Filter = "Excel File |*.xlsx",
                    Title = "Where do you want to save the site?"
                };
                file.ShowDialog();
                

                if (file.FileName?.Length != 0)
                {
                    using (var package = new ExcelPackage())
                    {
                        ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Sheet1");

                        worksheet.Cells[1, 1].Value = "Title";
                        worksheet.Cells[1, 2].Value = "Url";
                        worksheet.Cells[1, 3].Value = "AutoWriteLocation";
                        worksheet.Cells[1, 4].Value = "SaveReadLocation";
                        worksheet.Cells[1, 5].Value = "AutoStart";
                        worksheet.Cells[1, 6].Value = "Paths-->";

                        for (int i = 0; i < siteList.Count; i++)
                        {
                            worksheet.Cells[i + 2, 1].Value = siteList[i].Title;
                            worksheet.Cells[i + 2, 2].Value = siteList[i].Url;
                            worksheet.Cells[i + 2, 3].Value = siteList[i].AutoWriteLocation;
                            worksheet.Cells[i + 2, 4].Value = siteList[i].SaveLocation;
                            worksheet.Cells[i + 2, 5].Value = siteList[i].AutoStart ? "1" : "0";
                            for (int j = 0; j < siteList[i].Paths.Count; j++)
                            {
                                worksheet.Cells[i + 2, j + 6].Value = siteList[i].Paths[j].Description + pathSeperator +
                                                                  siteList[i].Paths[j].Function + pathSeperator +
                                                                  siteList[i].Paths[j].How + pathSeperator +
                                                                  siteList[i].Paths[j].Where;
                            }
                        }
                        
                        FileInfo excelFile = new FileInfo(file.FileName);
                        package.SaveAs(excelFile);

                        MessageBox.Show("Excel file saved in this location:\n" + file.FileName);
                    }

                }

            }
            else
            {
                MessageBox.Show("My sites list empty!");
            }
        }
        private void importButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog file = new OpenFileDialog
            {
                Filter = "Excel File |*.xlsx",
                Title = "Select the Excel file containing the site content."
            };
            file.ShowDialog();

            if (file.FileName?.Length != 0)
            {
                try
                {
                    List<Site> newSiteList = new List<Site>();

                    // Load the Excel file
                    FileInfo excelFile = new FileInfo(file.FileName);
                    using (var package = new ExcelPackage(excelFile))
                    {
                        // Get the first worksheet
                        ExcelWorksheet worksheet = package.Workbook.Worksheets.First<ExcelWorksheet>();


                        // Read the values from the worksheet into the list
                        int rowCount = worksheet.Dimension.Rows;
                        for (int i = 2; i <= rowCount; i++)
                        {
                            Site site = new Site();
                            site.Title = worksheet.Cells[i, 1].Value.ToString();
                            site.Url = worksheet.Cells[i, 2].Value.ToString();
                            site.AutoWriteLocation = worksheet.Cells[i, 3].Value.ToString();
                            site.SaveLocation = worksheet.Cells[i, 4].Value.ToString();
                            site.AutoStart = worksheet.Cells[i, 5].Value.ToString() == "1";

                            site.Paths = new List<PathItem>();
                            int j = 0;
                            while (worksheet.Cells[i, j + 6].Value != null)
                            {
                                string[] tempPath = worksheet.Cells[i, j + 6].Value.ToString().Split(pathSeperator);
                                PathItem path = new PathItem
                                {
                                    Description = tempPath[0],
                                    Function = tempPath[1],
                                    How = tempPath[2],
                                    Where = tempPath[3]
                                };

                                site.Paths.Add(path);
                                j++;
                            }

                            newSiteList.Add(site);
                        }

                        foreach (var newSite in newSiteList)
                        {
                            foreach (var title in siteList.Select(x => x.Title))
                            {
                                if (title == newSite.Title)
                                {
                                    newSite.Title += siteList.Count.ToString(); 
                                    MessageBox.Show("Warning!\nThere is another site with the same title:\n" + title +"\nTitle has updated:\n" + newSite.Title);
                                    break;
                                }

                            }
                            siteList.Add(newSite);
                        }

                        mySitesListBox.Items.Clear();

                        foreach (var mySite in siteList)
                        {
                            MySitesItem mySiteItem = new MySitesItem
                            {
                                Title = mySite.Title,
                                Url = mySite.Url,
                                AutoStart = mySite.AutoStart
                            };
                            mySitesListBox.Items.Add(mySiteItem);
                        }
                    }
                }
                catch (Exception)
                {
                    MessageBox.Show("The file layout is incorrect!");
                }

        }
        }
        #endregion


        #region BrowserGrid
        private void browserLocationButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog file = new OpenFileDialog
            {
                Filter = "Exe File |*.exe",
                Title = "Select the driver for browser."
            };
            file.ShowDialog();

            browserTextBlock.Text = file.FileName;
        }
        private void browserSaveButton_Click(object sender, RoutedEventArgs e)
        {
            if (browserTextBlock.Text?.Length > 0)
            {
                

                if (BrowserTest(browserComboBox.SelectedValue.ToString(), browserTextBlock.Text))
                {
                    browserLocation = browserTextBlock.Text;

                    browserName = browserComboBox.SelectedValue.ToString();

                    MessageBox.Show("Browser settings valid.");

                    browserSettingGrid.Visibility = Visibility.Hidden;
                }
                else
                {
                    MessageBox.Show("The selected values are incorrect, incompatible, or there is a version difference!");
                }
            }
            else
            {
                MessageBox.Show("Please choose driver location.");
            }
        }
        private bool BrowserTest(string name, string location)
        {
            
            try
            {
                if (name == "Microsoft Edge")
                {
                    EdgeOptions options = new EdgeOptions();
                    options.PageLoadStrategy = PageLoadStrategy.Eager; //Dont wait for images to load





                    driver = new EdgeDriver(location, options);
                }
                else if (name == "Google Chrome")
                {
                    ChromeOptions options = new ChromeOptions();
                    options.PageLoadStrategy = PageLoadStrategy.Eager; //Dont wait for images to load
                    driver = new ChromeDriver(location, options);
                }

                if (driver != null)
                {
                    driver.Url = "https://www.google.com";

                    driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(5);
                    driver.Manage().Window.Maximize();

                    driver.Quit();

                    return true;
                }


                return false;
            }
            catch (Exception)
            {
                MessageBox.Show("There is a problem with the browser or url.");
                return false;
            }
            
        }
        private void browserCancelButton_Click(object sender, RoutedEventArgs e)
        {
            browserSettingGrid.Visibility = Visibility.Hidden;
        }
        private void browserSettingsButton_Click(object sender, RoutedEventArgs e)
        {
            if (browserName == "Microsoft Edge")
            {
                browserComboBox.SelectedIndex = 0;
            }
            else if (browserName == "Google Chrome")
            {
                browserComboBox.SelectedIndex = 1;
            }
            else
            {
                browserComboBox.SelectedIndex = 0;
            }

            if (browserLocation != null)
            {
                browserTextBlock.Text = browserLocation;
            }
            else
            {
                browserTextBlock.Text = "";
            }
            
            browserSettingGrid.Visibility = Visibility.Visible;
        }
        #endregion
    }
}
