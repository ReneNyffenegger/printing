'
' [ https://github.com/ReneNyffenegger/runVBAFilesInOffice/blob/master/runVBAFilesInOffice.vbs ]
'
'    c:\lib\runVBAFilesInOffice\runVBAFilesInOffice.vbs -word a4-ruler -c main
'  

option explicit

dim ad_sh  as shapes

sub main()

    dim ps           as PageSetup: set ps = activeDocument.pageSetup

    dim l            as shape

    dim cm10         as double: cm10 = centimetersToPoints(10)
    dim pw           as double: pw   = pointsToCentimeters(ps.pageWidth )
    dim ph           as double: ph   = pointsToCentimeters(ps.pageHeight)

    set ad_sh = activeDocument.shapes


    ps.leftMargin   = 0
    ps.rightMargin  = 0
    ps.topMargin    = 0
    ps.bottomMargin = 0


  ' ---

    dim mg    as double: mg    = 1.5
    dim wg    as double: wg    = 2
    dim wg_mm as double: wg_mm = 0.2
    dim wg_c2 as double: wg_c2 = 0.3
    dim wg_cm as double: wg_cm = 0.5
    dim ln_mm as double: ln_mm = 0.3
    dim ln_c2 as double: ln_c2 = 0.4
    dim ln_cm as double: ln_cm = 0.6

    call line(     0,    mg,    pw-2*mg,     mg, wg)
    call line(    mg,  2*mg,    mg     ,     ph, wg)
    call line(  2*mg, ph-mg,         pw,  ph-mg, wg)
    call line( pw-mg,     0,    pw-mg  ,ph-2*mg, wg)

  ' ---

    dim i   as long
    dim lbl as shape

    for i = 0 to 10 * ( pw-2*mg ) ' {

        if     i mod 10 = 0 then

               call line(i/10, mg, i/10, mg+ln_cm, wg_cm)

               if i > 0 then
                  call text(i/10, i/10, mg+ln_cm, "h")
               end if

        elseif i mod  5 = 0 then

               call line(i/10, mg, i/10, mg+ln_c2, wg_c2)

        else

               call line(i/10, mg, i/10, mg+ln_mm, wg_mm)

        end if

    next i ' }

    dim n as long: n=0
    for i = 0 to 10 * (pw-2*mg) ' {

        if     i mod 10 = 0 then

               call line(pw-i/10, ph-mg, pw-i/10, ph-mg-ln_cm, wg_cm)

               if n > 0 then
                  call text(n, pw-i/10, ph-mg-ln_cm-1, "h")
               end if
               n = n+1

        elseif i mod  5 = 0 then

               call line(pw-i/10, ph-mg, pw-i/10, ph-mg-ln_c2, wg_c2)

        else

               call line(pw-i/10, ph-mg, pw-i/10, ph-mg-ln_mm, wg_mm)

        end if

    next i ' }

    n = 0
    for i = 0 to 10 * (ph - 2*mg) ' {

        if     i mod 10 = 0 then

               call line(pw-mg, i/10, pw-mg-ln_cm, i/10, wg_cm)
               
               if n > 0 then
                  call text(n, pw-mg-ln_cm-1, i/10, "v")
               end if
              
               n = n+1

        elseif i mod  5 = 0 then

               call line(pw-mg, i/10, pw-mg-ln_c2, i/10, wg_c2)

        else

               call line(pw-mg, i/10, pw-mg-ln_mm, i/10, wg_mm)

        end if

    next i ' }

    n = 0
    for i = 0 to 10 * (ph-2*mg) ' {

        if     i mod 10 = 0 then

               call line(mg, ph-i/10, mg+ln_cm, ph-i/10, wg_cm)

               if n > 0 then
                  call text(n, mg+ln_cm, ph-i/10, "v")
               end if
              
               n = n+1


        elseif i mod  5 = 0 then

               call line(mg, ph-i/10, mg+ln_c2, ph-i/10, wg_c2)

        else

               call line(mg, ph-i/10, mg+ln_mm, ph-i/10, wg_mm)

        end if

    next i ' }

    activeDocument.saved = true

end sub


private sub line(xs as double, ys as double, xe as double, ye as double, w as double) ' {

    dim line_ as shape

    set line_ = ad_sh.addLine(   _
       centimetersToPoints(xs), _
       centimetersToPoints(ys), _
       centimetersToPoints(xe), _
       centimetersToPoints(ye)  _
    )

    line_.line.weight = w * 2


end sub ' }

private sub text(txt as long, x as double, y as double, vert_or_horiz as string) ' {

   dim lbl as shape

   dim adj_v as double
   dim adj_h as double

   if vert_or_horiz = "h" then
      adj_h = -0.5
      adj_v =  0
   else
      adj_h =  0
      adj_v = -0.5
   end if
      

   set lbl = ad_sh.addLabel(msoTextOrientationHorizontal, centimetersToPoints(x+adj_h), centimetersToPoints(y+adj_v), centimetersToPoints(1), centimetersToPoints(1))

   lbl.textFrame.autoSize       = false
   lbl.textFrame.textRange.text = txt
   lbl.TextFrame.textRange.paragraphFormat.Alignment      = wdAlignParagraphCenter
   lbl.textFrame.textRange.paragraphFormat.spaceAfter     = 0

   lbl.textFrame.marginLeft     = centimetersToPoints(0)
   lbl.textFrame.marginRight    = centimetersToPoints(0)
   lbl.textFrame.marginTop      = centimetersToPoints(0)
   lbl.textFrame.marginBottom   = centimetersToPoints(0)

   lbl.textFrame.verticalAnchor = msoAnchorMiddle

   lbl.line.visible = false



 ' Why is this even necessary???
 
   doEvents
   lbl.width = centimetersToPoints(1)

end sub ' }
