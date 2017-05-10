//Copyright (c) 2008 Alexey Potapov

//Permission is hereby granted, free of charge, to any person obtaining a copy 
//of this software and associated documentation files (the "Software"), to deal
//in the Software without restriction, including without limitation the rights to 
//use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies 
//of the Software, and to permit persons to whom the Software is furnished to do 
//so, subject to the following conditions:

//The above copyright notice and this permission notice shall be included in all 
//copies or substantial portions of the Software.

//THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR 
//IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS 
//FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR 
//COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER 
//IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION 
//WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

using System;
using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;

namespace GreyableImage
{
  /// <summary>
  /// ImageGreyer class exposing attachable dependency property that when attached to an Image
  /// and set to true will couse that Image to turn greyscale when IsEnabled is set to false.
  /// 
  /// This is intended to be used for images in toolbars, menus or buttons where ability of an icon to 
  /// grey itself out when disabled is essential.
  /// 
  /// This class implements the attached property trick brilliantly described by Dan Crevier in his blog:
  ///    http://blogs.msdn.com/dancre/archive/2006/03/04/543854.aspx
  /// <remarks>
  /// 1) Greyscale image is created using FormatConvertedBitmap class. Unfortunately when converting the
  ///    image to greyscale this class does not preserve transparency information. To overcome that, there is 
  ///    an opacity mask created from original image that is applied to greyscale image in order to preserve
  ///    transparency information. Because of that if an OpacityMask is applied to original image that mask 
  ///    has to be combined with that special opacity mask of greyscale image in order to make a proper 
  ///    greyscale image look. If you know how to combine two opacity masks please let me know.
  /// 2) When specifying source Uri from XAML try to use Absolute Uri otherwise the greyscale image
  ///    may not be created in some scenarious. There is GetAbsoluteUri() method aiming to improve the situation 
  ///    by trying to generate an absolute Uri from given source, but I cannot guarantee it will work in all 
  ///    possible scenarious.
  /// 3) In case the greyscaled version cannot be created for whatever reason the original image with 
  ///    60% opacity (i.e. dull colours) will be used instead.
  /// 4) Changing Source from code will take precedence over Style triggers. Source set through triggers 
  ///    will be ignored once it was set from code. This is not the fault of the control, but is the way 
  ///    WPF works: http://msdn.microsoft.com/en-us/library/ms743230%28classic%29.aspx
  /// 5) Supports DrawingImage as a source, thanks to Morten Schou.
  /// </remarks>
  /// </summary>
  public class ImageGreyer
  {
    #region Fields

    // image this effect is attached to
    private Image _image;

    // these are holding references to original and greyscale ImageSources
    private ImageSource _sourceColour, _sourceGreyscale;

    // these are holding original and greyscale opacity masks
    private Brush _opacityMaskColour, _opacityMaskGreyscale;

    #endregion // Fields

    #region attachable properties

    #region IsGreyable

    /// <summary>
    /// Attach this property to standart WPF image and if set to true will make that image greyable
    /// </summary>
    public static DependencyProperty IsGreyableProperty =
      DependencyProperty.RegisterAttached("IsGreyable", typeof(bool), typeof(ImageGreyer),
                                          new PropertyMetadata(false, new PropertyChangedCallback(OnChangedIsGreyable)));

    /// <summary>
    /// Attached property accessors
    /// </summary>
    public static bool GetIsGreyable(DependencyObject sender)
    {
      return (bool)sender.GetValue(IsGreyableProperty);
    }
    public static void SetIsGreyable(DependencyObject sender, bool isGreyable)
    {
      sender.SetValue(IsGreyableProperty, isGreyable);
    }

    /// <summary>
    /// Callback when the IsGreyable property is set or changed.
    /// </summary>
    private static void OnChangedIsGreyable(DependencyObject dependencyObject, DependencyPropertyChangedEventArgs e)
    {
      Image image = dependencyObject as Image;
      if (null != image)
      {
        if ((bool)e.NewValue)
        {
          // turn greyability effect on if it is not turned on yet
          if (image.ReadLocalValue(GreyabilityEffectProperty) == DependencyProperty.UnsetValue)
          {
            ImageGreyer greyability = new ImageGreyer(image);
            image.SetValue(GreyabilityEffectProperty, greyability);
          }
        }
        else
        {
          // remove greyability effect
          if (image.ReadLocalValue(GreyabilityEffectProperty) != DependencyProperty.UnsetValue)
          {
            ImageGreyer greyability = (ImageGreyer)image.ReadLocalValue(GreyabilityEffectProperty);
            greyability.Detach();
            image.SetValue(GreyabilityEffectProperty, DependencyProperty.UnsetValue);
          }
        }
      }
    }

    #endregion // IsGreyable

    #region GreyabilityEffect

    /// <summary>
    /// attachable dependency property to be set on image to store reference to ourselves - private, used by this class only
    /// </summary>
    public static DependencyProperty GreyabilityEffectProperty =
      DependencyProperty.RegisterAttached("GreyabilityEffect", typeof(ImageGreyer), typeof(ImageGreyer));

    #endregion // GreyabilityEffect

    #endregion // attachable properties

    #region Constructor

    public ImageGreyer(Image image)
    {
      _image = image;

      // If the image is not initialized yet, the Source is not set and SetSource will return without caching
      // the sources. Still change notification for Source property will not be fired if the Source was set 
      // from XAML e.g. <Image Source="image.png"/>. In this case we have to wait until the Image is initialized
      // which will mean that the Source is set (if it is supposed to be set from XAML) and we can cache it.
      // otherwise we just call SetSources caching all requireв data.
      if (!_image.IsInitialized)
      {
        // delay attaching to an image untill it is ready
        _image.Initialized += OnChangedImageInitialized;
      }
      else
      {
        // attach greyability effect to an image
        Attach();
      }
    }

    #endregion // Constructor

    #region Event handlers

    /// <summary>
    /// Called when IsInitialized property of the Image is set to true
    /// </summary>
    void OnChangedImageInitialized(object sender, EventArgs e)
    {
      // image is ready for attaching greyability effect
      Attach();
    }

    /// <summary>
    /// Called when IsEnabled property of the Image is changed
    /// </summary>
    private void OnChangedImageIsEnabled(object sender, DependencyPropertyChangedEventArgs e)
    {
      UpdateImage();
    }

    /// <summary>
    /// Called when Source property of the Image is changed
    /// </summary>
    private void OnChangedImageSource(object sender, EventArgs e)
    {
      Image image = sender as Image;

      // only recache Source if it's a new one
      if (null != image &&
          !object.ReferenceEquals(image.Source, _sourceColour) &&
          !object.ReferenceEquals(image.Source, _sourceGreyscale))  
      {
        SetSources();

        // have to asynchronously invoke UpdateImage because it changes the Source property 
        // of an image, but we cannot change it from within its change notification handler.
        image.Dispatcher.BeginInvoke(DispatcherPriority.Background, new Action(UpdateImage));
      }
    }

    /// <summary>
    /// Called when OpacityMask property of the Image is changed
    /// </summary>
    private void OnChangedImageOpacityMask(object sender, EventArgs e)
    {
      Image image = sender as Image;

      // only recache opacityMask if it's a new one
      if (null != image &&
          !object.ReferenceEquals(image.OpacityMask, _opacityMaskColour) &&
          !object.ReferenceEquals(image.OpacityMask, _opacityMaskGreyscale))
      {
        _opacityMaskColour = image.OpacityMask;
      }
    }

    #endregion Event handlers

    #region Helper methods

    /// <summary>
    /// Attaching greyability effect to an Image
    /// </summary>
    private void Attach()
    {
      // first we need to cache original and greyscale Sources ...
      SetSources();
      
      // ... and OpacityMasks
      SetOpacityMasks();

      // now if the image is disabled we need to grey it out now
      UpdateImage();

      // set event handlers
      _image.IsEnabledChanged += OnChangedImageIsEnabled;

      // there is no change notification event for OpacityMask dependency property 
      // in Image class but we can use property descriptor to add value changed callback
      DependencyPropertyDescriptor dpDescriptor = DependencyPropertyDescriptor.FromProperty(Image.OpacityMaskProperty, typeof(Image));
      dpDescriptor.AddValueChanged(_image, OnChangedImageOpacityMask);

      // there is no change notification for Source dependency property
      // in Image class but we can use property descriptor to add value changed callback
      dpDescriptor = DependencyPropertyDescriptor.FromProperty(Image.SourceProperty, typeof(Image));
      dpDescriptor.AddValueChanged(_image, OnChangedImageSource);
    }

    /// <summary>
    /// Detaches this effect from the image, 
    /// </summary>
    private void Detach()
    {
      if (_image != null)
      {
        // remove all event handlers first...
        _image.IsEnabledChanged -= OnChangedImageIsEnabled;

        // there is no special change notification event for OpacityMask dependency property in Image class
        // but we can use property descriptor to remove value changed callback
        DependencyPropertyDescriptor dpDescriptor = DependencyPropertyDescriptor.FromProperty(Image.OpacityMaskProperty, typeof(Image));
        dpDescriptor.RemoveValueChanged(_image, OnChangedImageOpacityMask);

        // there is no change notification event for Source dependency property 
        // in Image class but we can use property descriptor to add value changed callback
        dpDescriptor = DependencyPropertyDescriptor.FromProperty(Image.SourceProperty, typeof(Image));
        dpDescriptor.RemoveValueChanged(_image, OnChangedImageSource);

        // in case the image is disabled we have to change the Source and OpacityMask 
        // properties back to the original values
        _image.Source = _sourceColour;
        _image.OpacityMask = _opacityMaskColour;

        // now release all the references we hold
        _image = null;
        _opacityMaskColour = _opacityMaskGreyscale = null;
        _sourceColour = _sourceGreyscale = null;
      }
    }

    /// <summary>
    /// Cashes original ImageSource, creates and caches greyscale ImageSource and greyscale opacity mask
    /// </summary>
    private void SetSources()
    {
      if (null == _image.Source)
        return;

      // in case greyscale image cannot be created set greyscale source to original Source first
      _sourceGreyscale = _sourceColour = _image.Source;

      try
      {
        BitmapSource colourBitmap;

        if (_sourceColour is DrawingImage)
        {
          // support for DrawingImage as a source - thanks to Morten Schou who provided this code
          colourBitmap = new RenderTargetBitmap((int)_sourceColour.Width,
                                                (int)_sourceColour.Height,
                                                96, 96,
                                                PixelFormats.Default);
          DrawingVisual drawingVisual = new DrawingVisual();
          DrawingContext drawingDC = drawingVisual.RenderOpen();

          drawingDC.DrawImage(_sourceColour,
                              new Rect(new Size(_sourceColour.Height,
                                                _sourceColour.Width)));
          drawingDC.Close();
          (colourBitmap as RenderTargetBitmap).Render(drawingVisual);
        }
        else
        {
          // get the string Uri for the original image source first
          String stringUri = TypeDescriptor.GetConverter(_sourceColour).ConvertTo(_sourceColour, typeof(string)) as string;

          // Create colour BitmapImage using an absolute Uri (generated from stringUri)
          colourBitmap = new BitmapImage(GetAbsoluteUri(stringUri));
        }

        // create and cache greyscale ImageSource
        _sourceGreyscale = new FormatConvertedBitmap(colourBitmap, PixelFormats.Gray8, null, 0);
      }
      catch (Exception e)
      {
        System.Diagnostics.Debug.Fail("The Image used cannot be greyed out.",
                                      "Make sure absolute Uri is used, relative Uri may sometimes resolve incorrectly.\n\nException: " + e.Message);
      }
    }

    /// <summary>
    /// Cashes original Image opacity mask, creates and caches greyscale Image opacity mask.
    /// </summary>
    private void SetOpacityMasks()
    {
      if (null == _image.Source)
        return;

      _opacityMaskColour = _image.OpacityMask;

      // create Opacity Mask for greyscale image as FormatConvertedBitmap used to 
      // create greyscale image does not preserve transparency info.
      _opacityMaskGreyscale = new ImageBrush(_sourceColour);
      _opacityMaskGreyscale.Opacity = 0.6;
    }

    /// <summary>
    /// Sets image source and opacity mask from cache.
    /// </summary>
    public void UpdateImage()
    {
      if (_image.IsEnabled)
      {
        // change Source and OpacityMask of an image back to original values
        _image.Source = _sourceColour;
        _image.OpacityMask = _opacityMaskColour;
      }
      else
      {
        // change Source and OpacityMask of an image to values generated for greyscale version
        _image.Source = _sourceGreyscale;
        _image.OpacityMask = _opacityMaskGreyscale;
      }
    }

    /// <summary>
    /// Creates and returns an absolute Uri using the path provided.
    /// Throws UriFormatException if an absolute URI cannot be created.
    /// </summary>
    /// <param name="stringUri">string uri</param>
    /// <returns>an absolute URI based on string URI provided</returns>
    /// <exception cref="UriFormatException - thrown when absolute Uri cannot be created from provided stringUri."/>
    /// <exception cref="ArgumentNullException - thrown when stringUri is null."/>
    private Uri GetAbsoluteUri(String stringUri)
    {
      Uri uri = null;

      // try to resolve it as an absolute Uri 
      // if uri is relative its likely to point in a wrong direction
      if (!Uri.TryCreate(stringUri, UriKind.Absolute, out uri))
      {
        // it seems that the Uri is relative, at this stage we can only assume that
        // the image requested is in the same assembly as this oblect,
        // so we modify the string Uri to make it absolute ...
        stringUri = "pack://application:,,,/" + stringUri.TrimStart(System.IO.Path.DirectorySeparatorChar, System.IO.Path.AltDirectorySeparatorChar);

        // ... and try to resolve again
        // at this stage if it doesn't resolve the UriFormatException is thrown
        uri = new Uri(stringUri);
      }

      return uri;
    }

    #endregion // Helper methods
  }
}

