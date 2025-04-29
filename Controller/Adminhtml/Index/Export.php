<?php

namespace PinBlooms\ProductImageExport\Controller\Adminhtml\Index;

use Magento\Backend\App\Action;
use Magento\Backend\App\Action\Context;
use Magento\Framework\App\Filesystem\DirectoryList;
use Magento\Framework\Controller\ResultFactory;
use Magento\Framework\App\Response\Http\FileFactory;
use Magento\Framework\Filesystem;
use Magento\Framework\Filesystem\Io\File;
use Magento\Catalog\Model\ResourceModel\Product\CollectionFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;
use Psr\Log\LoggerInterface;

class Export extends Action
{
    /**
     * @var FileFactory
     */
    protected $fileFactory;

    /**
     * @var Filesystem
     */
    protected $filesystem;

    /**
     * @var File
     */
    protected $file;

    /**
     * @var CollectionFactory
     */
    protected $productCollectionFactory;

    /**
     * @var LoggerInterface
     */
    protected $logger;

    /**
     * Constructor
     *
     * @param Context $context
     * @param FileFactory $fileFactory
     * @param Filesystem $filesystem
     * @param File $file
     * @param CollectionFactory $productCollectionFactory
     * @param LoggerInterface $logger
     */
    public function __construct(
        Context $context,
        FileFactory $fileFactory,
        Filesystem $filesystem,
        File $file,
        CollectionFactory $productCollectionFactory,
        LoggerInterface $logger
    ) {
        parent::__construct($context);
        $this->fileFactory = $fileFactory;
        $this->filesystem = $filesystem;
        $this->file = $file;
        $this->productCollectionFactory = $productCollectionFactory;
        $this->logger = $logger;
    }
    /**
     * Execute product export action
     *
     * @return \Magento\Framework\Controller\ResultInterface
     */
    public function execute()
    {
        $filePath = $this->exportProductData();
        if ($filePath) {
            $fileName = 'product_data_' . date('Y-m-d') . '.xlsx';
            return $this->fileFactory->create(
                $fileName,
                [
                    'type' => 'filename',
                    'value' => $filePath,
                    'rm' => true
                ],
                DirectoryList::VAR_DIR
            );
        } else {
            $this->messageManager->addErrorMessage(__('Failed to export product data.'));
        }
        $resultRedirect = $this->resultFactory->create(ResultFactory::TYPE_REDIRECT);
        return $resultRedirect->setPath('*/*/');
    }
    /**
     * Export selected products data to an Excel file.
     *
     * @return string|null
     */
    protected function exportProductData()
    {
        $selectedProductIds = $this->getRequest()->getParam('selected');
        if (empty($selectedProductIds)) {
            $this->messageManager->addErrorMessage(__('Please select products to export.'));
            return null;
        }
        $productCollection = $this->productCollectionFactory->create();
        $productCollection->addAttributeToSelect([
            'sku',
            'name',
            'price',
            'thumbnail',
            'short_description',
            'special_price',
            'color',
            'quantity_and_stock_status',
            'weight',
            'country_of_manufacture',
            'description'
        ]);
        $productCollection->addFieldToFilter('entity_id', ['in' => $selectedProductIds]);
        $productCollection->setPageSize(300);

        $countryFactory = $this->_objectManager->create(\Magento\Directory\Model\CountryFactory::class);
        $countryCollection = $countryFactory->create()->getCollection();

        $countryNames = [];
        foreach ($countryCollection as $country) {
            $countryNames[$country->getCountryId()] = $country->getName();
        }

        $stockRegistry = $this->_objectManager->get(\Magento\CatalogInventory\Api\StockRegistryInterface::class);

        if ($productCollection->getSize() > 0) {
            $spreadsheet = new Spreadsheet();
            $sheet = $spreadsheet->getActiveSheet();

            $sheet->setCellValue('A1', 'Sr No.');
            $sheet->setCellValue('B1', 'SKU');
            $sheet->setCellValue('C1', 'Name');
            $sheet->setCellValue('D1', 'Price');
            $sheet->setCellValue('E1', 'Image');
            $sheet->setCellValue('F1', 'Short Description');
            $sheet->setCellValue('G1', 'Special Price');
            $sheet->setCellValue('H1', 'Color');
            $sheet->setCellValue('I1', 'Qty');
            $sheet->setCellValue('J1', 'Availability');
            $sheet->setCellValue('K1', 'Weight(In gram)');
            $sheet->setCellValue('L1', 'Country of Manufacture');
            $sheet->setCellValue('M1', 'Description');

            $headingStyle = $sheet->getStyle('A1:M1');

            $headingStyle->getFont()->setBold(true);

            $row = 2;
            $serialNumber = 1;

            $imageWidth = 230;
            $imageHeight = 130;

            foreach ($productCollection as $product) {
                $countryCode = $product->getCountryOfManufacture();
                $countryName = isset($countryNames[$countryCode]) ? $countryNames[$countryCode] : '';

                $stockItem = $stockRegistry->getStockItem($product->getId());
                $availability = $stockItem->getIsInStock() ? 'In Stock' : 'Out of Stock';

                $shortDescription = strip_tags($product->getShortDescription() ?? '');
                $description = strip_tags($product->getDescription() ?? '');
                $mediaDirectory = $this->filesystem->getDirectoryRead(DirectoryList::MEDIA);
                $catalogProductPath = $mediaDirectory->getAbsolutePath('catalog/product');
                $thumbnailPath = $catalogProductPath . $product->getThumbnail();


                if (file_exists($thumbnailPath) && is_file($thumbnailPath)) {

                    $tempImagePath = $this->resizeAndCompressImage($thumbnailPath, $imageWidth, $imageHeight);

                    $sheet->setCellValue('A' . $row, $serialNumber);
                    $sheet->setCellValue('B' . $row, $product->getSku());
                    $sheet->setCellValue('C' . $row, $product->getName());
                    $sheet->setCellValue('D' . $row, $product->getPrice());
                    $sheet->setCellValue('E' . $row, '');
                    $sheet->setCellValue('F' . $row, $shortDescription);
                    $sheet->setCellValue('G' . $row, $product->getSpecialPrice());
                    $sheet->setCellValue('H' . $row, $product->getColor());
                    $sheet->setCellValue('I' . $row, $product->getQuantityAndStockStatus());
                    $sheet->setCellValue('J' . $row, $availability);
                    $sheet->setCellValue('K' . $row, $product->getWeight());
                    $sheet->setCellValue('L' . $row, $countryName);
                    $sheet->setCellValue('M' . $row, $description);

                    $drawing = new Drawing();
                    $drawing->setName('Product Image');
                    $drawing->setDescription('Product Image');
                    $drawing->setPath($tempImagePath);

                    $drawing->setCoordinates('E' . $row);
                    $drawing->setOffsetX(10);
                    $drawing->setOffsetY(18);
                    $drawing->setWidth($imageWidth);
                    $drawing->setHeight($imageHeight);
                    $drawing->setWorksheet($sheet);

                    $sheet->getRowDimension($row)->setRowHeight($imageHeight);

                    $serialNumber++;
                    $row++;
                } else {

                    $sheet->setCellValue('A' . $row, $serialNumber);
                    $sheet->setCellValue('B' . $row, $product->getSku());
                    $sheet->setCellValue('C' . $row, $product->getName());
                    $sheet->setCellValue('D' . $row, $product->getPrice());
                    $sheet->setCellValue('F' . $row, $product->getShortDescription());
                    $sheet->setCellValue('G' . $row, $product->getSpecialPrice());
                    $sheet->setCellValue('H' . $row, $product->getColor());
                    $sheet->setCellValue('I' . $row, $product->getQuantityAndStockStatus());
                    $sheet->setCellValue('J' . $row, '');
                    $sheet->setCellValue('K' . $row, $product->getWeight());
                    $sheet->setCellValue('L' . $row, $product->getCountryOfManufacture());
                    $sheet->setCellValue('M' . $row, $product->getDescription());
                    $row++;
                    $row++;

                    $serialNumber++;
                    $this->logger->error("Thumbnail image not found for product SKU: " . $product->getSku());
                }
            }

            $sheet->getColumnDimension('E')->setWidth($imageWidth / 8);

            $sheet->getColumnDimension('B')->setAutoSize(true);
            $sheet->getColumnDimension('C')->setAutoSize(true);
            $sheet->getColumnDimension('F')->setAutoSize(true);
            $sheet->getColumnDimension('M')->setAutoSize(true);
            $sheet->getColumnDimension('G')->setAutoSize(true);
            $sheet->getColumnDimension('H')->setAutoSize(true);
            $sheet->getColumnDimension('I')->setAutoSize(true);
            $sheet->getColumnDimension('J')->setAutoSize(true);
            $sheet->getColumnDimension('K')->setAutoSize(true);
            $sheet->getColumnDimension('L')->setAutoSize(true);

            $sheet->getStyle('B')->getAlignment()->setWrapText(true);
            $sheet->getStyle('C')->getAlignment()->setWrapText(true);
            $sheet->getStyle('F')->getAlignment()->setWrapText(true);
            $sheet->getStyle('M')->getAlignment()->setWrapText(true);
            $sheet->getStyle('G')->getAlignment()->setWrapText(true);
            $sheet->getStyle('H')->getAlignment()->setWrapText(true);
            $sheet->getStyle('I')->getAlignment()->setWrapText(true);
            $sheet->getStyle('J')->getAlignment()->setWrapText(true);
            $sheet->getStyle('K')->getAlignment()->setWrapText(true);
            $sheet->getStyle('L')->getAlignment()->setWrapText(true);

            $directoryWrite = $this->filesystem->getDirectoryWrite(DirectoryList::VAR_DIR);
            $filePath = $directoryWrite->getAbsolutePath('product_data.xlsx');
            $writer = new Xlsx($spreadsheet);
            $writer->save($filePath);

            return $filePath;
        } else {

            $this->messageManager->addErrorMessage(__('No products found for export.'));
            return null;
        }
    }

    /**
     * Resize and compress the image.
     *
     * @param string $imagePath The path to the original image file.
     * @param int $newWidth The desired width for the resized image.
     * @param int $newHeight The desired height for the resized image.
     * @return string The path to the resized and compressed image file.
     */

    protected function resizeAndCompressImage($imagePath, $newWidth, $newHeight)
    {
        if (!file_exists($imagePath) || !is_file($imagePath) || !getimagesize($imagePath)) {

            $this->logger->error("Invalid image file: $imagePath");
            return null;
        }

        $image = @imagecreatefromjpeg($imagePath);

        if (!$image) {
            $this->logger->error("Failed to create image from file: $imagePath");
            return null;
        }

        $originalWidth = imagesx($image);
        $originalHeight = imagesy($image);

        $newHeight = floor($originalHeight * ($newWidth / $originalWidth));

        $newImage = imagecreatetruecolor($newWidth, $newHeight);

        imagecopyresampled($newImage, $image, 0, 0, 0, 0, $newWidth, $newHeight, $originalWidth, $originalHeight);

        $tempImagePath = tempnam(sys_get_temp_dir(), 'product_image_');
        imagejpeg($newImage, $tempImagePath, 20);

        imagedestroy($image);
        imagedestroy($newImage);

        return $tempImagePath;
    }
}
