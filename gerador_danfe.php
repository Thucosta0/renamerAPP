<?php
// Gerador de DANFE - Integração com Python
// Verificar se autoloader existe (pode estar no diretório atual ou pai)
$autoload_paths = ['vendor/autoload.php', '../vendor/autoload.php'];
$autoload_found = false;

foreach ($autoload_paths as $path) {
    if (file_exists($path)) {
        require_once $path;
        $autoload_found = true;
        break;
    }
}

if (!$autoload_found) {
    echo "ERROR:Dependências PHP não instaladas! Execute: composer install\n";
    exit(1);
}

use NFePHP\DA\NFe\Danfe;

// Configurar codificação (sem mbstring)
ini_set('default_charset', 'UTF-8');

// Habilita exibição de erros para debug
error_reporting(E_ALL);
ini_set('display_errors', 1);
ini_set('memory_limit', '512M');  // Aumentar limite de memória

// Verifica se foi passado o arquivo XML como parâmetro
if ($argc < 2) {
    echo "ERROR:Arquivo XML não especificado!\n";
    echo "Uso: php gerador_danfe.php arquivo.xml [nome_personalizado]\n";
    exit(1);
}

$arquivoXML = $argv[1];
$nomePersonalizado = isset($argv[2]) ? $argv[2] : null; // Nome personalizado opcional

// Verifica se arquivo XML existe
if (!file_exists($arquivoXML)) {
    echo "ERROR:Arquivo '$arquivoXML' não encontrado!\n";
    exit(1);
}

try {
    // Carrega o XML da nota fiscal
    $xml = file_get_contents($arquivoXML);
    
    if ($xml === false) {
        throw new Exception("Não foi possível ler o arquivo XML");
    }
    
    if (empty($xml)) {
        throw new Exception("Arquivo XML está vazio");
    }
    
    // Verifica se é um XML válido
    $dom = new DOMDocument();
    libxml_use_internal_errors(true);
    $valid = $dom->loadXML($xml);
    
    if (!$valid) {
        $errors = libxml_get_errors();
        $errorMsg = "XML inválido: ";
        foreach ($errors as $error) {
            $errorMsg .= $error->message . " ";
        }
        throw new Exception($errorMsg);
    }
    
    // Verifica se tem as tags necessárias
    $xpath = new DOMXPath($dom);
    $infNFe = $xpath->query('//*[local-name()="infNFe"]');
    
    if ($infNFe->length == 0) {
        throw new Exception("Tag infNFe não encontrada no XML");
    }
    
    // Verifica modelo
    $modelo = $xpath->query('//*[local-name()="mod"]');
    if ($modelo->length > 0) {
        $mod = $modelo->item(0)->nodeValue;
        if ($mod != '55') {
            throw new Exception("Modelo deve ser 55 (NFe), encontrado: $mod");
        }
    } else {
        throw new Exception("Tag modelo não encontrada no XML");
    }
    
    // Cria o gerador de DANFE
    $danfe = new Danfe($xml);
    
    // Configurar formatação personalizada da chave de acesso
    // Formato: espaço no 9º dígito do final (################################### #########)
    $reflection = new ReflectionClass($danfe);
    $formatoChaveProperty = $reflection->getProperty('formatoChave');
    $formatoChaveProperty->setAccessible(true);
    $formatoChaveProperty->setValue($danfe, "################################### #########");
    
    // Ativar concatenação automática de informações sobre rastro
    $danfe->descProdInfoLoteTxt = true;
    
    // Gera o PDF da DANFE (formato A4 retrato - padrão Receita Federal)
    $pdf = $danfe->render();
    
    if (empty($pdf)) {
        throw new Exception("PDF gerado está vazio");
    }
    
    // Nome do arquivo PDF
    if ($nomePersonalizado) {
        // Se nome personalizado foi fornecido, usar ele (sem extensão)
        $nomeBase = basename($nomePersonalizado, '.pdf'); // Remove .pdf se fornecido
        $nomePDF = dirname($arquivoXML) . '/' . $nomeBase . '.pdf';
    } else {
        // Comportamento padrão - apenas o nome do XML
        $nomeBase = basename($arquivoXML, '.xml');
        $nomePDF = dirname($arquivoXML) . '/' . $nomeBase . '.pdf';
    }
    
    // Salva o PDF
    $bytesEscritos = file_put_contents($nomePDF, $pdf);
    
    if ($bytesEscritos === false) {
        throw new Exception("Não foi possível salvar o arquivo PDF");
    }
    
    if ($bytesEscritos == 0) {
        throw new Exception("PDF salvo está vazio");
    }
    
    // Verifica se o arquivo foi realmente criado
    if (!file_exists($nomePDF)) {
        throw new Exception("Arquivo PDF não foi criado");
    }
    
    // Retorna sucesso para o Python
    echo "SUCCESS:$nomePDF\n";
    
} catch (Exception $e) {
    // Retorna erro para o Python
    echo "ERROR:" . $e->getMessage() . "\n";
    exit(1);
}

/**
 * Formatar data do formato YYYY-MM-DD para DD/MM/YYYY
 */
function formatarData($data) {
    try {
        if (strlen($data) >= 10) {
            $ano = substr($data, 0, 4);
            $mes = substr($data, 5, 2);
            $dia = substr($data, 8, 2);
            return "$dia/$mes/$ano";
        }
        return $data;
    } catch (Exception $e) {
        return $data;
    }
}

?>
