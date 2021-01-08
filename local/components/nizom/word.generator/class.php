<?php
/**
 * Created by SpaceMedia.
 * Project Name: #
 * Developer Full Name: Najmiddinov Nizamiddin
 * Developer Mail: najmiddinov.nizom@gmail.com
 * Telegram: https://t.me/nizomiddin_n
 * Resume page: https://dev.1c-bitrix.ru/learning/resume.php?ID=36871542-1260669
 */

if (!defined('B_PROLOG_INCLUDED') || B_PROLOG_INCLUDED !== true) die();

use \Bitrix\Main\{ Loader, IO };
use \Bitrix\DocumentGenerator;

Loader::includeModule('main');
Loader::includeModule('form');
Loader::includeModule('documentgenerator');
Loader::includeModule('highloadblock');

class SpaceMediaWordGeneratorBaseComponent extends \CBitrixComponent
{
    const DOCUMENT_PATH = "/local/lens/request.docx",
          DOCUMENT_TEMPLATE =  __DIR__ . '/template.docx';

    private static $relativeFields = [],
                   $relativeData = [];

    public function onPrepareComponentParams($params)
    {
        return $params;
    }

    /**
     * Starting Form entity
     **/
    private static function getFormFields($entity)
    {
        $fields = [];
        $rsProperties = \CFormField::GetList($entity, "ALL", $by = "s_id", $orderField = "asc", $arFilter, $is_filtered);

        while ($arProp = $rsProperties->Fetch())
        {
            $fields[$arProp['SID']] = ['TYPE' => 'TEXT'];
        }

        return $fields;
    }

    private static function getFormData($elementId)
    {
        $request = \Bitrix\Main\Application::getInstance()->getContext()->getRequest();

        $values = [];

        $arAnswer = \CFormResult::GetDataByID($elementId, [], $arResult2, $arAnswer2);

        foreach ($arAnswer as $sid => $answer)
        {
            if(count($answer) > 1)
            {
                $val = implode(', ', array_diff(array_column($answer, 'USER_TEXT'), ["", false]));
            }
            else
            {
                $answer = current($answer);

                if($answer['FIELD_TYPE'] == 'checkbox')
                {
                    $val = 'Да';
                }
                elseif($answer['FIELD_TYPE'] == 'file')
                {
                    $val = ($request->isHttps() ? 'https://' : 'http://') . $request->getHttpHost() . \CFile::GetPath($answer['USER_FILE_ID']);
                }
                else
                {
                    $val = $answer['USER_TEXT'];
                }
            }

            $values[$sid] = $val;
        }

        return $values;
    }

    /**
     * Starting Highload entity
     **/
    private static function compileEntity($entity)
    {
        $hlblock = \Bitrix\Highloadblock\HighloadBlockTable::getById($entity)->fetch();
        $hlEntity = \Bitrix\Highloadblock\HighloadBlockTable::compileEntity($hlblock);

        self::$relativeFields = $hlEntity->getFields();
        self::$relativeData = $hlEntity->getDataClass();
    }

    private static function getRelativesFieldsEntity()
    {
        return self::$relativeFields;
    }

    private static function getRelativesDataEntity()
    {
        return self::$relativeData;
    }

    private static function getRelativesFields($entityId)
    {
        $result = [];

        foreach (self::getRelativesFieldsEntity() as $field){
            $name = str_replace(
                '_', '',
                ucwords(strtolower($field->getName()), '_')
            );

            $result["TableItem{$name}"] = ['VALUE' => "Table.Item.{$name}"];
        }

        return $result;
    }

    private static function getRelatives($entityId, $requestId)
    {
        $result = $relatives = [];

        foreach (self::getRelativesFieldsEntity() as $field)
        {
            $name = str_replace(
                '_', '',
                ucwords(strtolower($field->getName()), '_')
            );

            $result["TableItem{$name}"] = "Table.Item.{$name}";
        }


        $data = self::getRelativesDataEntity()::getList([
            'filter' => [
                'UF_REQUEST_ID' => $requestId
            ]
        ])->fetchAll();

        foreach ($data as $index => $relative)
        {
            foreach ($relative as $code => $value)
            {
                $name = str_replace(
                    '_', '',
                    ucwords(strtolower($code), '_')
                );
                $relatives[$index][$name] = $value;
            }
        }

        $result = array_merge([
            'Table' => new DocumentGenerator\DataProvider\ArrayDataProvider(
                $relatives,
                [
                    'ITEM_NAME' => 'Item',
                    'ITEM_PROVIDER' => DocumentGenerator\DataProvider\HashDataProvider::class,
                ]
            )
        ], $result);

        return $result;
    }

    /**
     * Merge Data
     **/
    private static function compileDocument($formEntity, $relativeEntity, $elementId)
    {
        self::compileEntity($relativeEntity);

        $formFields = self::getFormFields($formEntity);
        $relativeFields = self::getRelativesFields($relativeEntity);
        $fields = array_merge($formFields, $relativeFields);


        $formData = self::getFormData($elementId);
        $dataRelatives = self::getRelatives($relativeEntity, $elementId);
        $data = array_merge($dataRelatives, $formData);

        $file = new IO\File(self::DOCUMENT_TEMPLATE);
        $body = new DocumentGenerator\Body\Docx($file->getContents());

        $body->normalizeContent();
        $body->setValues($data);
        $body->setFields($fields);
        $result = $body->process();

        if(!$result->isSuccess())
        {
            $result = [
                'error' => $result->getErrorMessages()
            ];
        }
        else
        {
            file_put_contents($_SERVER['DOCUMENT_ROOT'].self::DOCUMENT_PATH, $body->getContent());

            $result = [
                'url' => self::DOCUMENT_PATH
            ];
        }

        return $result;
    }

    public function executeComponent()
    {
        $result = self::compileDocument(
            $this->arParams['ENTITY_ID'],
            $this->arParams['ENTITY_RELATIVE'],
            $this->arParams['ELEMENT_ID']
        );

        return $result['url'];
    }
}