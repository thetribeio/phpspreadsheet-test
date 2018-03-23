<?php

namespace AppBundle\Form\Type;

use Symfony\Component\Form\AbstractType;
use Symfony\Component\Form\Extension\Core\Type\ChoiceType;
use Symfony\Component\Form\FormBuilderInterface;
use Symfony\Component\Validator\Constraints as Assert;

class ExcelFormatType extends AbstractType
{
    public function buildForm(FormBuilderInterface $builder, array $options)
    {
        $builder
            ->add('format', ChoiceType::class, [
                'choices' => [
                    'xlsx' => 'xlsx',
                    'ods' => 'ods',
                    'csv' => 'csv',
                ],
                'label' => false,
                'placeholder' => 'Select a format',
            ]);
    }
}
