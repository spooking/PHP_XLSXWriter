<?php

require_once 'XLSXWriter.class.php';

class XLSXWriterHelper
{
    public static function writeToStdOut($filename = '导出', $config = [], $ds = [])
    {
        ob_end_clean();
        ob_start();
        header('Content-Disposition:attachment;filename='.$filename.'.xlsx');
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Transfer-Encoding: binary');
        header('Cache-Control: must-revalidate');
        header('Pragma: no-cache');
        header('Expires: 0');
        cookie(MODULE_NAME.'_'.CONTROLLER_NAME.'_exportExcel', ''.time());
        $writer = new \XLSXWriter();
        $writer->setTempDir('./Uploadfile/temp');
        $style = ['border' => 'left,right,top,bottom', 'border-style' => 'thin', 'valign' => 'center', 'wrap_text' => 'true'];

        $hf = [
            '序号' => 'string',
        ]; //列格式
        $hw = [8]; //列宽度
        $hs = [$style]; //列样式

        foreach ($config as $v) {
            if (count($v) >= 2) {
                $hs[] = $style;
                $hf[$v[1]] = count($v) >= 4 ? $v[3] : 'string';
                $hw[] = count($v) >= 3 ? $v[2] : 12;
            }
        }

        $hs['widths'] = $hw;

        $writer->writeSheetHeader('Sheet1', $hf, $hs);
        $data = [];
        $i = 0;
        foreach ($ds as $d) {
            $line = [++$i];

            foreach ($config as $v) {
                if (count($v) >= 1) {
                    $line[] = $d[$v[0]];
                }
            }
            $data[] = $line;
        }
        foreach ($data as $d) {
            $writer->writeSheetRow('Sheet1', $d, $style);
        }
        $writer->writeToStdOut();
        die;
    }

    private function getHdsWidths($hds)
    {
        $re = [];
        foreach ($hds as $hid => $h) {
            foreach ($h as $lid => $l) {
                if (array_key_exists(1, $l) && (int) $l[1] > 0) {
                    $re[$lid] = (int) $l[1];
                }
            }
        }

        return $re;
    }

    private function getHdsRowFs($hds)
    {
        $re = [];
        foreach ($hds as $hid => $h) {
            foreach ($h as $lid => $l) {
                if (array_key_exists(2, $l)) {
                    $re[$lid] = $l[2];
                }
            }
        }

        return $re;
    }

    private function getHdsRowDs($hds)
    {
        $re = [];
        $lm = 0;
        foreach ($hds as $hid => $h) {
            $ls = max(array_keys($h));
            if ($ls > $lm) {
                $lm = $ls;
            }
        }
        foreach ($hds as $hid => $h) {
            $re[$hid] = array_pad([], $lm, '');
            foreach ($h as $lid => $l) {
                $re[$hid][$lid] = $l[0];
            }
        }

        return $re;
    }

    private function getHdsMergeInfo($hds)
    {
        $re = [];
        foreach ($hds as $hid => $h) {
            foreach ($h as $lid => $l) {
                if (array_key_exists('merge', $l)) {
                    $re[] = $l['merge'];
                } else {
                    if (!empty($l[0]) && $hid < (count($hds) - 1) && $hds[$hid + 1][$lid][0] == '') {
                        $re[] = [$hid, $lid, $hid + 1, $lid];
                    }
                }
            }
        }

        return $re;
    }

    private function config2HDS($config, &$hds, $h = 0, $l = 0)
    {
        $hs = [];
        $i = -1;
        foreach ($config as $v) {
            ++$i;
            if (!array_key_exists($h, $hds)) {
                $hds[$h] = [];
            }

            if ($v[0] == '-') {
                $w = self::config2HDS($v[2], $hds, $h + 1, $l + $i);
                $hds[$h][$l + $i] = [$v[1], 'merge' => [$h, $l + $i, $h, $l + $i + $w]];
                for ($k = 1; $k <= $w; ++$k) {
                    $hds[$h][$l + $i + $k] = [''];
                }
                $i += $w;
            } else {
                $hds[$h][$l + $i] = [$v[1], $v[2], $v[0]];
            }
        }

        return $i;
    }

    public static function writeToStdOutX($filename = '导出', $config = [], $ds = [])
    {
        ob_end_clean();
        ob_start();
        header('Content-Disposition:attachment;filename='.$filename.'.xlsx');
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Transfer-Encoding: binary');
        header('Cache-Control: must-revalidate');
        header('Pragma: no-cache');
        header('Expires: 0');
        cookie(MODULE_NAME.'_'.CONTROLLER_NAME.'_exportExcel', ''.time());
        $writer = new \XLSXWriter();
        $writer->setTempDir('./Uploadfile/temp');
        $style = ['border' => 'left,right,top,bottom', 'border-style' => 'thin', 'valign' => 'center', 'wrap_text' => 'true', 'height' => 20];

        $hds = [];
        self::config2HDS($config, $hds);
        $hd = self::getHdsRowDs($hds);
        $hw = self::getHdsWidths($hds);
        $hf = self::getHdsRowFs($hds);
        $hm = self::getHdsMergeInfo($hds);
        $hs = [
            'suppress_row' => true,
            'widths' => $hw,
        ];
        $writer->writeSheetHeader('Sheet1', array_pad([], count($hw), 'string'), $col_options = $hs);
        foreach ($hd as $dhi => $dh) {
            $writer->writeSheetRow('Sheet1', $dh, array_merge($style, $dhi == 0 ? ['halign' => 'center',  'font-style' => 'bold'] : ['halign' => 'center']));
        }

        $data = [];
        $i = 0;

        foreach ($ds as $d) {
            $line = [++$i];

            foreach ($hf as $hk => $v) {
                $line[$hk] = $d[$v];
            }
            $data[] = $line;
        }

        foreach ($data as $d) {
            $writer->writeSheetRow('Sheet1', $d, $style);
        }
        foreach ($hm as $mg) {
            $writer->markMergedCell('Sheet1', $start_row = $mg[0], $start_col = $mg[1], $end_row = $mg[2], $end_col = $mg[3]);
        }
        $writer->writeToStdOut();
        die;
    }
}
